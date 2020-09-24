VERSION 5.00
Object = "{C151518A-D64D-4C66-96F3-DB69BF286B30}#1.0#0"; "WinXPC Engine.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Azkary"
   ClientHeight    =   6990
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "PopupControl Sample"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":058A
   ScaleHeight     =   6990
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin Azkary.ucTimer ucTimerRotateAzkar 
      Left            =   7380
      Top             =   6840
      _ExtentX        =   450
      _ExtentY        =   423
      Interval        =   1
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000001&
      Height          =   1030
      Left            =   1530
      ScaleHeight     =   975
      ScaleWidth      =   1155
      TabIndex        =   64
      Top             =   3645
      Width           =   1220
      Begin VB.PictureBox picPosition 
         Height          =   240
         Index           =   8
         Left            =   855
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   73
         Top             =   675
         Width           =   240
      End
      Begin VB.PictureBox picPosition 
         Height          =   240
         Index           =   7
         Left            =   450
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   72
         Top             =   675
         Width           =   240
      End
      Begin VB.PictureBox picPosition 
         Height          =   240
         Index           =   6
         Left            =   45
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   71
         Top             =   675
         Width           =   240
      End
      Begin VB.PictureBox picPosition 
         Height          =   240
         Index           =   5
         Left            =   855
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   70
         Top             =   360
         Width           =   240
      End
      Begin VB.PictureBox picPosition 
         Height          =   240
         Index           =   4
         Left            =   450
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   69
         Top             =   360
         Width           =   240
      End
      Begin VB.PictureBox picPosition 
         Height          =   240
         Index           =   3
         Left            =   45
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   68
         Top             =   360
         Width           =   240
      End
      Begin VB.PictureBox picPosition 
         Height          =   240
         Index           =   2
         Left            =   855
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   67
         Top             =   45
         Width           =   240
      End
      Begin VB.PictureBox picPosition 
         Height          =   240
         Index           =   1
         Left            =   450
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   66
         Top             =   45
         Width           =   240
      End
      Begin VB.PictureBox picPosition 
         Height          =   240
         Index           =   0
         Left            =   45
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   65
         Top             =   45
         Width           =   240
      End
   End
   Begin VB.ComboBox cboAlignment 
      Height          =   315
      Left            =   5940
      TabIndex        =   63
      Top             =   4815
      Width           =   1680
   End
   Begin VB.TextBox txtShowDelay 
      Height          =   285
      Left            =   5940
      TabIndex        =   60
      Top             =   3960
      Width           =   1005
   End
   Begin VB.TextBox txtPeriod 
      Height          =   285
      Left            =   1530
      TabIndex        =   58
      Top             =   2745
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   9060
      Left            =   8550
      Picture         =   "frmMain.frx":8046
      ScaleHeight     =   9000
      ScaleWidth      =   12000
      TabIndex        =   47
      Top             =   0
      Width           =   12060
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   285
         Left            =   270
         TabIndex        =   74
         Top             =   2835
         Width           =   1320
      End
      Begin VB.ComboBox cboLanguage 
         Height          =   315
         Left            =   180
         TabIndex        =   57
         Top             =   5625
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   285
         Left            =   225
         TabIndex        =   55
         Top             =   4635
         Width           =   1320
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   285
         Left            =   225
         TabIndex        =   54
         Top             =   4365
         Width           =   1320
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   285
         Left            =   225
         TabIndex        =   53
         Top             =   4095
         Width           =   1320
      End
      Begin VB.CommandButton cmdAzkar 
         Caption         =   "Azkar"
         Height          =   285
         Left            =   270
         TabIndex        =   52
         Top             =   2565
         Width           =   1320
      End
      Begin VB.CommandButton cmdShowZekr 
         Caption         =   "Preview"
         Height          =   285
         Left            =   225
         TabIndex        =   51
         Top             =   1665
         Width           =   1320
      End
      Begin VB.CommandButton cmdSaveZekr 
         Caption         =   "Save"
         Height          =   285
         Left            =   225
         TabIndex        =   50
         Top             =   1395
         Width           =   1320
      End
      Begin VB.CommandButton cmdDeleteZekr 
         Caption         =   "Delete"
         Height          =   285
         Left            =   225
         TabIndex        =   49
         Top             =   405
         Width           =   1320
      End
      Begin VB.CommandButton cmdAddZekr 
         Caption         =   "Add New"
         Height          =   285
         Left            =   225
         TabIndex        =   48
         Top             =   135
         Width           =   1320
      End
      Begin VB.Label lblLanguage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Language"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   56
         Top             =   5400
         Width           =   1410
      End
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   6705
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox picBalloonBackColor 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2340
      ScaleHeight     =   225
      ScaleWidth      =   315
      TabIndex        =   46
      Top             =   5715
      Width           =   375
   End
   Begin VB.PictureBox picBalloonColor 
      BackColor       =   &H00000080&
      Height          =   285
      Left            =   1845
      ScaleHeight     =   225
      ScaleWidth      =   315
      TabIndex        =   45
      Top             =   5715
      Width           =   375
   End
   Begin VB.CheckBox chkShowBalloon 
      Height          =   190
      Left            =   1530
      TabIndex        =   44
      Top             =   5715
      Width           =   190
   End
   Begin Azkary.ToolTipOnDemand ToolTipOnDemand 
      Left            =   6255
      Top             =   6840
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin Azkary.TrayIcon TrayIcon 
      Left            =   5760
      Top             =   6840
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin MSComctlLib.Slider sldTransparency 
      Height          =   285
      Left            =   5940
      TabIndex        =   42
      Top             =   3555
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   503
      _Version        =   393216
      Max             =   255
      TickFrequency   =   15
   End
   Begin MSComctlLib.Slider sldAnimationTime 
      Height          =   285
      Left            =   5940
      TabIndex        =   41
      Top             =   3150
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   503
      _Version        =   393216
      LargeChange     =   50
      Max             =   5000
      TickFrequency   =   200
   End
   Begin VB.CommandButton cmdPaste 
      Height          =   330
      Left            =   6435
      Picture         =   "frmMain.frx":F99D
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1125
      Width           =   420
   End
   Begin VB.CommandButton cmdCopy 
      Height          =   330
      Left            =   6030
      Picture         =   "frmMain.frx":1005F
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1125
      Width           =   420
   End
   Begin VB.CommandButton cmdCut 
      Height          =   330
      Left            =   5625
      Picture         =   "frmMain.frx":10721
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1125
      Width           =   420
   End
   Begin VB.CommandButton cmdFontUnderline 
      Height          =   330
      Left            =   5220
      Picture         =   "frmMain.frx":1086B
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1125
      Width           =   420
   End
   Begin VB.CommandButton cmdFontItalic 
      Height          =   330
      Left            =   4815
      Picture         =   "frmMain.frx":10A35
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1125
      Width           =   420
   End
   Begin VB.CommandButton cmdFontBold 
      Height          =   330
      Left            =   4410
      Picture         =   "frmMain.frx":10BFF
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1125
      Width           =   420
   End
   Begin VB.ComboBox cboFontSize 
      Height          =   315
      Left            =   3240
      TabIndex        =   34
      Top             =   1125
      Width           =   780
   End
   Begin VB.ComboBox cboFontName 
      Height          =   315
      Left            =   765
      Sorted          =   -1  'True
      TabIndex        =   33
      Top             =   1125
      Width           =   2490
   End
   Begin VB.CommandButton cmdColor 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4005
      Picture         =   "frmMain.frx":10D49
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1125
      Width           =   420
   End
   Begin RichTextLib.RichTextBox txtZekr 
      Height          =   1185
      Left            =   765
      TabIndex        =   31
      Top             =   1440
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   2090
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":110D7
   End
   Begin VB.PictureBox picTheme 
      FillColor       =   &H000080FF&
      Height          =   330
      Left            =   5940
      ScaleHeight     =   270
      ScaleWidth      =   1620
      TabIndex        =   30
      Top             =   4365
      Width           =   1680
   End
   Begin MSComctlLib.ImageList imglTheme 
      Left            =   3015
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.ComboBox cboPopupHeight 
      Height          =   315
      Left            =   2880
      TabIndex        =   29
      Top             =   3195
      Width           =   870
   End
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3780
      Top             =   6840
   End
   Begin VB.CommandButton cmdRemoveSoundFile 
      Height          =   285
      Left            =   7560
      Picture         =   "frmMain.frx":11167
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6435
      Width           =   420
   End
   Begin VB.CommandButton cmdStopSound 
      Height          =   330
      Left            =   4365
      Picture         =   "frmMain.frx":114E2
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6120
      Width           =   420
   End
   Begin Azkary.ucTimer ucTimer 
      Index           =   0
      Left            =   5355
      Top             =   6840
      _ExtentX        =   450
      _ExtentY        =   423
      Interval        =   1000
   End
   Begin VB.CommandButton cmdPlaySound 
      Height          =   330
      Left            =   4860
      Picture         =   "frmMain.frx":1184C
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton cmdRecordSound 
      Height          =   330
      Left            =   3915
      Picture         =   "frmMain.frx":11BA5
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6120
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   4815
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MCI.MMControl MMC 
      Height          =   330
      Left            =   -1125
      TabIndex        =   23
      Top             =   6930
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdBrowseForFile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7110
      Picture         =   "frmMain.frx":11FA1
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6435
      Width           =   420
   End
   Begin VB.TextBox txtSoundFile 
      Height          =   285
      Left            =   1530
      TabIndex        =   21
      Top             =   6480
      Width           =   5550
   End
   Begin VB.CheckBox chkPlaySound 
      Height          =   240
      Left            =   1530
      TabIndex        =   19
      Top             =   6165
      Width           =   240
   End
   Begin VB.CheckBox chkShowPopup 
      Height          =   190
      Left            =   1530
      TabIndex        =   17
      Top             =   5355
      Width           =   190
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -1215
      Top             =   6840
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   4
      Common_Dialog   =   0   'False
      PictureControl  =   0   'False
   End
   Begin VB.CheckBox chkEnabled 
      Height          =   190
      Left            =   1530
      TabIndex        =   14
      Top             =   4995
      Width           =   190
   End
   Begin VB.ComboBox cboPopupWidth 
      Height          =   315
      Left            =   1530
      TabIndex        =   11
      Top             =   3195
      Width           =   645
   End
   Begin VB.ListBox lstAzkar 
      Height          =   1035
      ItemData        =   "frmMain.frx":1238A
      Left            =   765
      List            =   "frmMain.frx":1238C
      TabIndex        =   7
      Top             =   45
      Width           =   7530
   End
   Begin VB.PictureBox picNotify 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   4365
      Picture         =   "frmMain.frx":1238E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ComboBox cboAnimation 
      Height          =   315
      Left            =   5940
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2745
      Width           =   2340
   End
   Begin VB.Label lblAlignment 
      BackStyle       =   0  'Transparent
      Caption         =   "Alignment:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4410
      TabIndex        =   62
      Top             =   4770
      Width           =   1320
   End
   Begin VB.Label lblSeconds 
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6975
      TabIndex        =   61
      Top             =   3960
      Width           =   1275
   End
   Begin VB.Label lblMinutes 
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2565
      TabIndex        =   59
      Top             =   2745
      Width           =   1320
   End
   Begin VB.Label lblShowTray 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Balloon:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   270
      TabIndex        =   43
      Top             =   5670
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   3555
      Picture         =   "frmMain.frx":12918
      Top             =   6165
      Width           =   240
   End
   Begin VB.Label lblPopupHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2250
      TabIndex        =   28
      Top             =   3195
      Width           =   555
   End
   Begin VB.Label lblSoundFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Sound File:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      TabIndex        =   20
      Top             =   6525
      Width           =   825
   End
   Begin VB.Label lblPlaySound 
      BackStyle       =   0  'Transparent
      Caption         =   "Play Sound:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      TabIndex        =   18
      Top             =   6165
      Width           =   870
   End
   Begin VB.Label lblShowPopup 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Popup:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      TabIndex        =   16
      Top             =   5310
      Width           =   960
   End
   Begin VB.Label lblAzkar 
      BackStyle       =   0  'Transparent
      Caption         =   "Azkar:"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   225
      TabIndex        =   15
      Top             =   135
      Width           =   645
   End
   Begin VB.Label lblEnabled 
      BackStyle       =   0  'Transparent
      Caption         =   "Enabled:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      TabIndex        =   13
      Top             =   4995
      Width           =   645
   End
   Begin VB.Label lblPopupPosition 
      BackStyle       =   0  'Transparent
      Caption         =   "Popup Position:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      TabIndex        =   12
      Top             =   3645
      Width           =   1230
   End
   Begin VB.Label lblPopupWidth 
      BackStyle       =   0  'Transparent
      Caption         =   "Popup Width:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      TabIndex        =   10
      Top             =   3195
      Width           =   1095
   End
   Begin VB.Label lblPeriod 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Every:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      TabIndex        =   9
      Top             =   2745
      Width           =   960
   End
   Begin VB.Label lblZekr 
      BackStyle       =   0  'Transparent
      Caption         =   "Zekr:"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   225
      TabIndex        =   8
      Top             =   1350
      Width           =   420
   End
   Begin VB.Label lblTransparency 
      BackStyle       =   0  'Transparent
      Caption         =   "Transparency:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4365
      TabIndex        =   5
      Top             =   3555
      Width           =   1335
   End
   Begin VB.Label lblShowTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Time:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4365
      TabIndex        =   4
      Top             =   3960
      Width           =   1395
   End
   Begin VB.Label lblAnimationTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Animation Time:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4365
      TabIndex        =   3
      Top             =   3150
      Width           =   1350
   End
   Begin VB.Label lblTheme 
      BackStyle       =   0  'Transparent
      Caption         =   "Theme:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4365
      TabIndex        =   2
      Top             =   4365
      Width           =   1335
   End
   Begin VB.Label lblAnimation 
      BackStyle       =   0  'Transparent
      Caption         =   "Animation:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4365
      TabIndex        =   0
      Top             =   2790
      Width           =   1275
   End
   Begin VB.Menu mnuTrayPopup 
      Caption         =   "TrayPopup"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuTrayShow 
         Caption         =   "Show Window"
      End
      Begin VB.Menu mnuTrayHide 
         Caption         =   "Hide Window"
      End
      Begin VB.Menu mnuTray1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTryDisable 
         Caption         =   "&Disable"
      End
      Begin VB.Menu mnuTray2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayAzkar 
         Caption         =   "Azkar"
      End
      Begin VB.Menu mnuTray3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'====================================================================
Dim WithEvents Tray As CTray
Attribute Tray.VB_VarHelpID = -1

'====================================================================
Dim CurrentPopup As Integer
Dim StopPlaying As Boolean
Dim bLoadingLanguage As Boolean

Private Sub Form_Load()
    
    Dim x As Long
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    Me.Left = GetSettings(AppRegPath, "Settings", "MainLeft", 1000)
    Me.Top = GetSettings(AppRegPath, "Settings", "MainTop", 1000)
    'Me.Width = GetSettings(AppRegPath, "Settings", "MainWidth", 11685)
    'Me.Height = GetSettings(AppRegPath, "Settings", "MainHeight", 8085)
    '----------------------------------------------------------------
    LabelsColor = GetSettings(AppRegPath, "Settings", "LabelsColor", vbWhite)
    AutoStartUp = GetSettings(AppRegPath, "Settings", "AutoStartUp", 1)
    RotateAzkar = GetSettings(AppRegPath, "Settings", "RotateAzkar", 0)
    RotateAzkarPeriod = GetSettings(AppRegPath, "Settings", "RotateAzkarPeriod", 2.15)
    AzkarStatus = GetSettings(AppRegPath, "Settings", "AzkarStatus", 1)
    '----------------------------------------------------------------
    bLoaded = False
    RunningStatus = False
    
    AzkarFile = AppPath + "Azkar.ini"
    LanguageFile = AppPath + "Language.ini"
    LoadLanguage
    bLoaded = False
    '----------------------------------------------------------------
    Me.Caption = Me.Caption
    PrepareFiles
    '----------------------------------------------------------------
    ' Get the system metrics we need
    ScreenWidth = GetSystemMetrics(SM_CXFULLSCREEN) * Screen.TwipsPerPixelX
    ScreenHeight = GetSystemMetrics(SM_CYFULLSCREEN) * Screen.TwipsPerPixelY
    'lngScaleX = Me.Width - Me.ScaleWidth
    'lngScaleY = Me.Height - Me.ScaleHeight
    'Debug.Print Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, ScreenWidth, ScreenHeight
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    PreparePopupSize 2
    PrepareAnimationList 0
    PrepareFontLists
    PrepareAlignment
    
    txtZekr.Font.Name = cboFontName.Text
    txtZekr.Font.Size = cboFontSize.Text
    cmdColor.BackColor = vbBlack
    cboPopupWidth.ListIndex = 7
    cboPopupHeight.ListIndex = 5
    cboAnimation.ListIndex = 0
    txtPeriod.Text = 2
    
    bLoaded = False
    LoadAzkarFile
    
    bLoaded = False
    
    If lstAzkar.ListCount > 0 Then
        lstAzkar.ListIndex = 0
        EditZekr 0
    End If
    '----------------------------------------------------------------
    cmdStopSound.Enabled = False
    '----------------------------------------------------------------
    '================================================================
    If AutoStartUp = 1 Then
        AutoRun = eAlways
    Else
        AutoRun = eNever
    End If
    '================================================================
    With TrayIcon
      .TrayIconVisible = True
      .IconHandle = Me.Icon
      .ToolTip = Me.Caption
      .Create Me.hwnd
      ''lblComctlVersion.Caption = .CommonControlsVersion
      ''lblSysTrayHWnd.Caption = .SysTrayHWnd
    End With
    '================================================================
    '----------------------------------------------------------------
    WindowsXPC1.ColorScheme = XP_OliveGreen '= System ' = XP_Blue
    WindowsXPC1.InitSubClassing
    '----------------------------------------------------------------
    RotateAzkarIndex = 0
    If AzkarStatus = 1 Then
        StartAzkars
    End If
    '----------------------------------------------------------------
    bLoaded = True
    'ShowWindow Me.hWnd, SW_SHOWNOACTIVATE
    'MakeTransparent Me
    Dim Ctl  As Control
    For Each Ctl In fMainForm.Controls
        If TypeOf Ctl Is Label Then
            Ctl.ForeColor = LabelsColor
        End If
    Next
    
    tmrHide.Enabled = True
    
End Sub

Private Sub cmdOptions_Click()
    frmOptions.Show vbModal, Me
    CenterForm frmOptions, Me
End Sub

Sub PrepareAlignment()
    Dim lIndex As Long
    lIndex = cboAlignment.ListIndex
    If lIndex < 0 Then lIndex = 1
    cboAlignment.Clear
    cboAlignment.AddItem AlignmentLeft
    cboAlignment.AddItem AlignmentRight
    cboAlignment.AddItem AlignmentCenter
    cboAlignment.ListIndex = lIndex
End Sub

Private Sub cmdAzkar_Click()
    frmBrowser.Show vbModal, Me
End Sub

Private Sub Form_Resize()
    Exit Sub
End Sub

Private Sub mnuTrayAzkar_Click()
    frmBrowser.Show 0 ', Me
End Sub

Sub PrepareFontLists()

    Dim x As Long
    
    '------------------------------------------------------
    'Load system fonts
    For x = 0 To Screen.FontCount - 1
        cboFontName.AddItem Screen.Fonts(x)
    Next
    cboFontName.Text = "Times New Roman"
    '------------------------------------------------------
    For x = 8 To 12
        cboFontSize.AddItem x
    Next
    cboFontSize.AddItem 14
    cboFontSize.AddItem 16
    cboFontSize.AddItem 18
    cboFontSize.AddItem 20
    cboFontSize.AddItem 22
    cboFontSize.AddItem 24
    cboFontSize.AddItem 26
    cboFontSize.AddItem 28
    cboFontSize.AddItem 36
    cboFontSize.AddItem 72
    cboFontSize.Text = 22
End Sub

Private Sub cboLanguage_Click()
    Dim Lang As String
    
    If bLoadingLanguage = True Then Exit Sub
    If bLoaded = False Then Exit Sub
    
    Lang = cboLanguage.List(cboLanguage.ListIndex)
    Lang = Trim(Lang)
    
    If Lang = "" Then Exit Sub
    
    WriteINI "Settings", "DefaultLanguage", Lang, LanguageFile
    LoadLanguage
    
    PrepareAlignment
    
    Me.Refresh

End Sub

Sub LoadLanguage()
     
    Dim Langs() As String
    Dim x As Long
    Dim Lang As String
    Dim Sections()  As String
        
    bLoadingLanguage = True
    
    LanguageFile = AppPath + "Language.ini"
    
    DefaultLanguage = ReadINI("Settings", "DefaultLanguage", LanguageFile)
    Langs = GetSectionsINI(LanguageFile)
    
    cboLanguage.Clear
    
    For x = LBound(Langs) To UBound(Langs)
        Lang = Trim(Langs(x))
        If Lang <> "" And Lang <> "Settings" Then
                cboLanguage.AddItem Lang
                If Lang = DefaultLanguage Then
                    cboLanguage.ListIndex = cboLanguage.ListCount - 1
                End If
        End If
    Next
    
    LanguageDirection = ReadINI(DefaultLanguage, "Direction", LanguageFile)
    
    Dim Ctl As Control
    
    Me.Caption = ReadINI(DefaultLanguage, "Title", LanguageFile)
    App.Title = Me.Caption
    
    lblAzkar.Caption = ReadINI(DefaultLanguage, "lblAzkar", LanguageFile)
    lblZekr.Caption = ReadINI(DefaultLanguage, "lblZekr", LanguageFile)
    lblPeriod.Caption = ReadINI(DefaultLanguage, "lblPeriod", LanguageFile)
    lblPopupWidth.Caption = ReadINI(DefaultLanguage, "lblPopupWidth", LanguageFile)
    lblPopupHeight.Caption = ReadINI(DefaultLanguage, "lblPopupHeight", LanguageFile)
    lblPopupPosition.Caption = ReadINI(DefaultLanguage, "lblPopupPosition", LanguageFile)
    lblEnabled.Caption = ReadINI(DefaultLanguage, "lblEnabled", LanguageFile)
    lblShowPopup.Caption = ReadINI(DefaultLanguage, "lblShowPopup", LanguageFile)
    
    lblShowTray.Caption = ReadINI(DefaultLanguage, "lblShowTray", LanguageFile)
    
    lblPlaySound.Caption = ReadINI(DefaultLanguage, "lblPlaySound", LanguageFile)
    lblSoundFile.Caption = ReadINI(DefaultLanguage, "lblSoundFile", LanguageFile)
    lblAnimation.Caption = ReadINI(DefaultLanguage, "lblAnimation", LanguageFile)
    lblAnimationTime.Caption = ReadINI(DefaultLanguage, "lblAnimationTime", LanguageFile)
    lblShowTime.Caption = ReadINI(DefaultLanguage, "lblShowTime", LanguageFile)
    lblTransparency.Caption = ReadINI(DefaultLanguage, "lblTransparency", LanguageFile)
    lblTheme.Caption = ReadINI(DefaultLanguage, "lblTheme", LanguageFile)
    lblAlignment.Caption = ReadINI(DefaultLanguage, "lblAlignment", LanguageFile)
    cmdAddZekr.Caption = ReadINI(DefaultLanguage, "cmdAddZekr", LanguageFile)
    cmdDeleteZekr.Caption = ReadINI(DefaultLanguage, "cmdDeleteZekr", LanguageFile)
    cmdSaveZekr.Caption = ReadINI(DefaultLanguage, "cmdSaveZekr", LanguageFile)
    cmdShowZekr.Caption = ReadINI(DefaultLanguage, "cmdShowZekr", LanguageFile)
    cmdAzkar.Caption = ReadINI(DefaultLanguage, "cmdAzkar", LanguageFile)
    cmdClose.Caption = ReadINI(DefaultLanguage, "cmdClose", LanguageFile)
    cmdAbout.Caption = ReadINI(DefaultLanguage, "cmdAbout", LanguageFile)
    cmdExit.Caption = ReadINI(DefaultLanguage, "cmdExit", LanguageFile)
    cmdBrowseForFile.Caption = ReadINI(DefaultLanguage, "cmdBrowseForFile", LanguageFile)
    cmdRemoveSoundFile.Caption = ReadINI(DefaultLanguage, "cmdRemoveSoundFile", LanguageFile)
    cmdRecordSound.Caption = ReadINI(DefaultLanguage, "cmdRecordSound", LanguageFile)
    cmdStopSound.Caption = ReadINI(DefaultLanguage, "cmdStopSound", LanguageFile)
    cmdPlaySound.Caption = ReadINI(DefaultLanguage, "cmdPlaySound", LanguageFile)
    cmdOptions.Caption = ReadINI(DefaultLanguage, "cmdOptions", LanguageFile)
    
    lblLanguage.Caption = ReadINI(DefaultLanguage, "lblLanguage", LanguageFile)
    MinutesLabel = ReadINI(DefaultLanguage, "lblMinutes", LanguageFile)
    lblMinutes.Caption = MinutesLabel
    lblSeconds.Caption = ReadINI(DefaultLanguage, "lblSeconds", LanguageFile)
    
    RotateAzkarLabel = ReadINI(DefaultLanguage, "RotateAzkarLabel", LanguageFile)
    RotationTimeLabel = ReadINI(DefaultLanguage, "RotationTimeLabel", LanguageFile)
    LabelsColorLabel = ReadINI(DefaultLanguage, "LabelsColorLabel", LanguageFile)
    AutoStartMessage = ReadINI(DefaultLanguage, "AutoStartMessage", LanguageFile)
    AzkarStatusLabel = ReadINI(DefaultLanguage, "AzkarStatusLabel", LanguageFile)
    
    AlignmentLeft = ReadINI(DefaultLanguage, "AlignmentLeft", LanguageFile)
    AlignmentRight = ReadINI(DefaultLanguage, "AlignmentRight", LanguageFile)
    AlignmentCenter = ReadINI(DefaultLanguage, "AlignmentCenter", LanguageFile)
    
    MenuTrayEnable = ReadINI(DefaultLanguage, "MenuTrayEnable", LanguageFile)
    MenuTrayDisable = ReadINI(DefaultLanguage, "MenuTrayDisable", LanguageFile)
    MenuTrayShow = ReadINI(DefaultLanguage, "MenuTrayShow", LanguageFile)
    MenuTrayHide = ReadINI(DefaultLanguage, "MenuTrayHide", LanguageFile)
    MenuTrayExit = ReadINI(DefaultLanguage, "MenuTrayExit", LanguageFile)
    MenuTrayAzkar = ReadINI(DefaultLanguage, "MenuTrayAzkar", LanguageFile)
    
    ExitMessage = ReadINI(DefaultLanguage, "ExitMessage", LanguageFile)
    
    CounterLabel = ReadINI(DefaultLanguage, "lblCounterLabel", LanguageFile)
    ResetCounter = ReadINI(DefaultLanguage, "cmdResetCounter", LanguageFile)
    CounterUp = ReadINI(DefaultLanguage, "cmdCounterUp", LanguageFile)
    Counter100Label = ReadINI(DefaultLanguage, "lblCounter100Label", LanguageFile)
    CounterReset100 = ReadINI(DefaultLanguage, "cmdCounterReset100", LanguageFile)
    CounterUp100 = ReadINI(DefaultLanguage, "cmdCounterUp100", LanguageFile)
    AzkarIndex = ReadINI(DefaultLanguage, "AzkarIndex", LanguageFile)
    AzkarClose = ReadINI(DefaultLanguage, "AzkarClose", LanguageFile)
    
    CommandOK = ReadINI(DefaultLanguage, "CommandOK", LanguageFile)
    CommandCancel = ReadINI(DefaultLanguage, "CommandCancel", LanguageFile)
   
    If RunningStatus = True Then
        mnuTryDisable.Caption = MenuTrayDisable
    Else
        mnuTryDisable.Caption = MenuTrayEnable
    End If
    
    mnuTrayShow.Caption = MenuTrayShow
    mnuTrayHide.Caption = MenuTrayHide
    mnuTrayExit.Caption = MenuTrayExit
    mnuTrayAzkar.Caption = MenuTrayAzkar
    
    If LanguageDirection = "RTL" Then
        Me.RightToLeft = True
        lstAzkar.RightToLeft = True
        
        SetParaDirection txtZekr.hwnd, PFE_RTLPAR
        'txtZekr.RightToLeft = True
        'txtZekr.Alignment = vbRightJustify
    Else
        Me.RightToLeft = False
        lstAzkar.RightToLeft = False
        SetParaDirection txtZekr.hwnd, Not PFE_RTLPAR
        'txtZekr.RightToLeft = False
        'txtZekr.Alignment = vbLeftJustify
    End If
    
    On Error Resume Next
    For Each Ctl In Me.Controls
        'TypeOf ctl Is Label
'            If TypeOf ctl Is TextBox _
'                Or TypeOf ctl Is ListBox Then
'                ctl.RightToLeft = True
'            End If
        Ctl.Refresh
    Next
    
    Me.Refresh
    bLoadingLanguage = False
    
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal, Me
    CenterForm frmAbout, Me
End Sub

Private Sub LoadAzkarFile()
    
    Dim Counter As Long
    Dim SelectedIndex As Long
    Dim x As Long
    Dim Section As String
    Dim Sections() As String
    Dim SectionNum As Long
                
    SelectedIndex = lstAzkar.ListIndex
    If SelectedIndex < 0 Then SelectedIndex = 0
    If SelectedIndex > lstAzkar.ListCount Then SelectedIndex = lstAzkar.ListCount - 1
    lstAzkar.Clear
    
    Counter = 0
     
    Sections = GetSectionsINI(AzkarFile)
            
    If SafeUBound(Sections) >= 0 Then
        ReDim Azkars(SafeUBound(Sections)) As Azkars
    Else
        Exit Sub
    End If
    
    For x = SafeLBound(Sections) To SafeUBound(Sections)
        
        Section = Sections(x)
        SectionNum = CLng(Val(Replace(Section, "Zekr-", "")))
        
        Azkars(Counter).ID = SectionNum
        
        Azkars(Counter).Zekr = ReadINI(Section, "Zekr", AzkarFile)
        
        lstAzkar.AddItem CStr(Counter + 1) + "- " + Azkars(Counter).Zekr
        
        Azkars(Counter).Period = Val(ReadINI(Section, "Period", AzkarFile))
        Azkars(Counter).PopupWidth = Val(ReadINI(Section, "PopupWidth", AzkarFile))
        Azkars(Counter).PopupHeight = Val(ReadINI(Section, "PopupHeight", AzkarFile))
        Azkars(Counter).PopupPosition = Val(ReadINI(Section, "PopupPosition", AzkarFile))
        Azkars(Counter).Enabled = Val(ReadINI(Section, "Enabled", AzkarFile))
        Azkars(Counter).ShowPopup = Val(ReadINI(Section, "ShowPopup", AzkarFile))
        
        Azkars(Counter).ShowBalloon = Val(ReadINI(Section, "ShowBalloon", AzkarFile))
        Azkars(Counter).BalloonColor = Val(ReadINI(Section, "BalloonColor", AzkarFile))
        Azkars(Counter).BalloonBackColor = Val(ReadINI(Section, "BalloonBackColor", AzkarFile))
        
        Azkars(Counter).PlaySound = Val(ReadINI(Section, "PlaySound", AzkarFile))
        Azkars(Counter).Animation = Val(ReadINI(Section, "Animation", AzkarFile))
        Azkars(Counter).AnimationTime = Val(ReadINI(Section, "AnimationTime", AzkarFile))
        Azkars(Counter).ShowDelay = Val(ReadINI(Section, "ShowDelay", AzkarFile))
        Azkars(Counter).Transparency = Val(ReadINI(Section, "Transparency", AzkarFile))
        Azkars(Counter).Theme = ReadINI(Section, "Theme", AzkarFile)
        Azkars(Counter).SoundFile = ReadINI(Section, "SoundFile", AzkarFile)
        
        Azkars(Counter).FontName = ReadINI(Section, "FontName", AzkarFile)
        Azkars(Counter).FontSize = Val(ReadINI(Section, "FontSize", AzkarFile))
        Azkars(Counter).FontColor = Val(ReadINI(Section, "FontColor", AzkarFile))
        Azkars(Counter).FontBold = Val(ReadINI(Section, "FontBold", AzkarFile))
        Azkars(Counter).FontItalic = Val(ReadINI(Section, "FontItalic", AzkarFile))
        Azkars(Counter).FontUnderline = Val(ReadINI(Section, "FontUnderline", AzkarFile))
        
        Counter = Counter + 1
    Next
    
    AzkarCount = Counter
    
    If Counter > 0 Then
        'lstAzkar.ListIndex = SelectedIndex
    End If
    
End Sub

Private Sub cmdAddZekr_Click()
    
    Dim Counter As Long
    Dim bOldLoaded As Boolean
        
    bOldLoaded = bLoaded
    bLoaded = False
    
    Counter = lstAzkar.ListCount
    
    lstAzkar.AddItem CStr(Counter + 1) + "-)New Azkar " + CStr(Counter + 1)
    
    lstAzkar.ListIndex = Counter
    
    txtZekr.Text = "Enter your azkar text here # " + CStr(Counter + 1)
    
    chkEnabled.Value = 1
    chkShowPopup.Value = 1
    chkShowBalloon.Value = 0
    chkPlaySound.Value = 0
    
    txtSoundFile.Text = ""
        
    Set picTheme.Picture = LoadPicture(AppPath + "themes\" + "Default.jpg")
    picTheme.Tag = "Default.jpg"
    picTheme.Refresh
    
    ReDim Preserve Azkars(Counter) As Azkars
    Azkars(Counter).ID = Counter
    
    cmdSaveZekr_Click
    
    StopAzkars
    StartAzkars
    
    bLoaded = bOldLoaded
End Sub

Private Sub cmdSaveZekr_Click()
    
    If txtZekr.Text = "" Then
        MsgBox "Please enter Azkar text", vbExclamation Or vbOKOnly, "Error"
        txtZekr.SetFocus
        Exit Sub
    End If
        
    Dim TextLine As String
    Dim Counter As Long
    Dim FileNum As Long
    Dim Section As String
    Dim TotalCount As Long
    Dim x As Long
    
    TotalCount = lstAzkar.ListCount
    Counter = lstAzkar.ListIndex
    
    Section = "Zekr-" + CStr(lstAzkar.ListIndex)
    
    If Counter < 0 Then Counter = 0
    '----------------------------------------------------------------
    txtSoundFile.Text = Trim(txtSoundFile.Text)
    
    If chkPlaySound.Value And txtSoundFile.Text = "" Then
        chkPlaySound.Value = 0
    End If
    
    If txtSoundFile.Text <> "" And FileExists(txtSoundFile.Text) = False Then
        chkPlaySound.Value = 0
    End If
    '----------------------------------------------------------------
    txtZekr.Text = Trim(txtZekr.Text)
    cboFontName.Text = Trim(cboFontName.Text)
    cboFontSize.Text = Trim(cboFontSize.Text)
    If cboFontName.Text = "" Then cboFontName.Text = "Times New Roman"
    If cboFontSize.Text = "" Then cboFontSize.Text = "22"
    
    'cmdAbout.Caption = ReadINI(DefaultLanguage, "cmdAbout", LanguageFile)
    
    WriteINI Section, "Zekr", Trim(txtZekr.Text), AzkarFile
    
    WriteINI Section, "Period", CStr(Val(txtPeriod.Text)), AzkarFile
    
    WriteINI Section, "PopupWidth", CStr(cboPopupWidth.ListIndex), AzkarFile
    WriteINI Section, "PopupHeight", CStr(cboPopupHeight.ListIndex), AzkarFile
    
    For x = 0 To 8
        If picPosition(x).BackColor = vbBlue Then
            WriteINI Section, "PopupPosition", CStr(x), AzkarFile
            Exit For
        End If
    Next
    
    WriteINI Section, "Enabled", CStr(chkEnabled.Value), AzkarFile
    WriteINI Section, "ShowPopup", CStr(chkShowPopup.Value), AzkarFile
    
    WriteINI Section, "ShowBalloon", CStr(chkShowBalloon.Value), AzkarFile
    WriteINI Section, "BalloonColor", CStr(picBalloonColor.BackColor), AzkarFile
    WriteINI Section, "BalloonBackColor", CStr(picBalloonBackColor.BackColor), AzkarFile
    
    WriteINI Section, "PlaySound", CStr(chkPlaySound.Value), AzkarFile
    WriteINI Section, "Animation", CStr(cboAnimation.ListIndex), AzkarFile
    WriteINI Section, "AnimationTime", CStr(sldAnimationTime.Value), AzkarFile
    
    WriteINI Section, "ShowDelay", CStr(Val(txtShowDelay.Text)), AzkarFile
    
    WriteINI Section, "Alignment", CStr(Val(cboAlignment.ListIndex)), AzkarFile
    
    WriteINI Section, "Transparency", CStr(sldTransparency.Value), AzkarFile
    WriteINI Section, "Theme", picTheme.Tag, AzkarFile
    WriteINI Section, "SoundFile", txtSoundFile.Text, AzkarFile
    
    WriteINI Section, "FontName", cboFontName.Text, AzkarFile
    WriteINI Section, "FontSize", cboFontSize.Text, AzkarFile
    WriteINI Section, "FontColor", CStr(cmdColor.BackColor), AzkarFile
    
    WriteINI Section, "FontBold", CStr(txtZekr.Font.Bold), AzkarFile
    WriteINI Section, "FontItalic", CStr(txtZekr.Font.Italic), AzkarFile
    WriteINI Section, "FontUnderline", CStr(txtZekr.Font.Underline), AzkarFile
    
    '----------------------------------------------------------------
    lstAzkar.List(lstAzkar.ListIndex) = CStr(lstAzkar.ListIndex) + "-" + txtZekr.Text
    '----------------------------------------------------------------
    Azkars(Counter).Zekr = txtZekr.Text
    Azkars(Counter).Period = Val(txtPeriod.Text)
    Azkars(Counter).PopupWidth = cboPopupWidth.ListIndex
    Azkars(Counter).PopupHeight = cboPopupHeight.ListIndex
    
    For x = 0 To 8
        If picPosition(x).BackColor = vbBlue Then
            Azkars(Counter).PopupPosition = x
            Exit For
        End If
    Next
    
    Azkars(Counter).Enabled = chkEnabled.Value
    Azkars(Counter).ShowPopup = chkShowPopup.Value
    
    Azkars(Counter).ShowBalloon = chkShowBalloon.Value
    Azkars(Counter).BalloonColor = picBalloonColor.BackColor
    Azkars(Counter).BalloonBackColor = picBalloonBackColor.BackColor
    
    Azkars(Counter).PlaySound = chkPlaySound.Value
    Azkars(Counter).Animation = cboAnimation.ListIndex
    Azkars(Counter).AnimationTime = sldAnimationTime.Value
    Azkars(Counter).ShowDelay = Val(txtShowDelay.Text)
    Azkars(Counter).Transparency = sldTransparency.Value
    Azkars(Counter).Theme = picTheme.Tag
    Azkars(Counter).Alignment = cboAlignment.ListIndex
    
    Azkars(Counter).SoundFile = txtSoundFile.Text
    
    Azkars(Counter).FontName = cboFontName.Text
    Azkars(Counter).FontSize = cboFontSize.Text
    Azkars(Counter).FontColor = cmdColor.BackColor
    Azkars(Counter).FontBold = txtZekr.Font.Bold
    Azkars(Counter).FontItalic = txtZekr.Font.Italic
    Azkars(Counter).FontUnderline = txtZekr.Font.Underline
    '----------------------------------------------------------------
    'LoadAzkarFile
    EditZekr lstAzkar.ListIndex
    '----------------------------------------------------------------
    StopAzkars
    StartAzkars
   
End Sub

Sub EditZekr(Optional ByVal iIndex As Long = 0)
    
    Dim Section As String
    
    Section = "Zekr-" + CStr(iIndex)
     
    txtZekr.Text = ReadINI(Section, "Zekr", AzkarFile)
    If txtZekr.Text = "" Then Exit Sub
        
    'cboPeriod.ListIndex = Val(ReadINI(Section, "Period", AzkarFile))
    txtPeriod.Text = Val(ReadINI(Section, "Period", AzkarFile))
    
    cboPopupWidth.ListIndex = Val(ReadINI(Section, "PopupWidth", AzkarFile))
    cboPopupHeight.ListIndex = Val(ReadINI(Section, "PopupHeight", AzkarFile))
    
    picPosition_Click Val(ReadINI(Section, "PopupPosition", AzkarFile))
    
    chkEnabled.Value = Val(ReadINI(Section, "Enabled", AzkarFile))
    chkShowPopup.Value = Val(ReadINI(Section, "ShowPopup", AzkarFile))
    
    chkShowBalloon.Value = Val(ReadINI(Section, "ShowBalloon", AzkarFile))
    picBalloonColor.BackColor = Val(ReadINI(Section, "BalloonColor", AzkarFile))
    picBalloonBackColor.BackColor = Val(ReadINI(Section, "BalloonBackColor", AzkarFile))
    
    chkPlaySound.Value = Val(ReadINI(Section, "PlaySound", AzkarFile))
    cboAnimation.ListIndex = Val(ReadINI(Section, "Animation", AzkarFile))
    sldAnimationTime.Value = Val(ReadINI(Section, "AnimationTime", AzkarFile))
    
    txtShowDelay.Text = Val(ReadINI(Section, "ShowDelay", AzkarFile))
    
    cboAlignment.ListIndex = Val(ReadINI(Section, "Alignment", AzkarFile))

    sldTransparency.Value = Val(ReadINI(Section, "Transparency", AzkarFile))
    picTheme.Tag = ReadINI(Section, "Theme", AzkarFile)
    txtSoundFile.Text = ReadINI(Section, "SoundFile", AzkarFile)
               
    cboFontName.Text = ReadINI(Section, "FontName", AzkarFile)
    txtZekr.Font.Name = cboFontName.Text
    
    cboFontSize.Text = ReadINI(Section, "FontSize", AzkarFile)
    txtZekr.Font.Size = cboFontSize.Text
    
    txtZekr.SelStart = 0
    txtZekr.SelLength = Len(txtZekr.Text)
    txtZekr.Font.Bold = CBool(ReadINI(Section, "FontBold", AzkarFile))
    txtZekr.Font.Italic = CBool(ReadINI(Section, "FontItalic", AzkarFile))
    txtZekr.Font.Underline = CBool(ReadINI(Section, "FontUnderline", AzkarFile))
        
    cmdColor.BackColor = ReadINI(Section, "FontColor", AzkarFile)
    'txtZekr.SelColor = cmdColor.BackColor
               
    picBalloonColor.BackColor = Val(ReadINI(Section, "BalloonColor", AzkarFile))
    picBalloonBackColor.BackColor = Val(ReadINI(Section, "BalloonBackColor", AzkarFile))
               
    If picTheme.Tag <> "" Then
        If FileExists(AppPath + "themes\" + picTheme.Tag) = True Then
            Set picTheme.Picture = LoadPicture(AppPath + "themes\" + picTheme.Tag)
            picTheme.Refresh
        Else
            picTheme.Tag = ""
            Set picTheme.Picture = Nothing
            picTheme.Refresh
        End If
    Else
        picTheme.Tag = ""
        Set picTheme.Picture = Nothing
        picTheme.Refresh
    End If
    
    If chkPlaySound.Value Then
        txtSoundFile.Enabled = True
        txtSoundFile.Refresh
        cmdBrowseForFile.Enabled = True
        cmdRemoveSoundFile.Enabled = True
        cmdRecordSound.Enabled = True
        cmdPlaySound.Enabled = True
    Else
        txtSoundFile.Enabled = False
        cmdBrowseForFile.Enabled = False
        cmdRemoveSoundFile.Enabled = False
        cmdRecordSound.Enabled = False
        cmdPlaySound.Enabled = False
        txtSoundFile.Refresh
    End If
                
End Sub

Private Sub cmdDeleteZekr_Click()
    Dim ret As Long
    Dim msgDelete As String
    Dim iIndex As Long
    Dim Counter  As Long
    Dim Section As String
    Dim x As Long
    
    If lstAzkar.ListIndex < 0 Then Exit Sub
    iIndex = lstAzkar.ListIndex
    
    msgDelete = ReadINI(DefaultLanguage, "DeleteMessage", LanguageFile)
    ret = MsgBox(msgDelete, vbExclamation Or vbYesNoCancel)
    
    If ret = vbYes Then
        StopAzkars
 
        'DeleteSectionINI "Zekr-" + CStr(iIndex), AzkarFile
        If FileExists(AzkarFile + ".bak") = True Then
            Kill AzkarFile + ".bak"
        End If
        
        Name AzkarFile As AzkarFile + ".bak"
        
        Counter = 0
        For x = LBound(Azkars) To UBound(Azkars)
            If x <> iIndex Then
                Section = "Zekr-" + CStr(Counter)
                WriteINI Section, "Zekr", Azkars(x).Zekr, AzkarFile
                WriteINI Section, "Period", Azkars(x).Period, AzkarFile
                WriteINI Section, "PopupWidth", Azkars(x).PopupWidth, AzkarFile
                WriteINI Section, "PopupHeight", Azkars(x).PopupHeight, AzkarFile
                WriteINI Section, "PopupPosition", Azkars(x).PopupPosition, AzkarFile
                WriteINI Section, "Enabled", Azkars(x).Enabled, AzkarFile
                WriteINI Section, "ShowPopup", Azkars(x).ShowPopup, AzkarFile
                
                WriteINI Section, "ShowBalloon", Azkars(x).ShowBalloon, AzkarFile
                WriteINI Section, "BalloonColor", Azkars(x).BalloonColor, AzkarFile
                WriteINI Section, "BalloonBackColor", Azkars(x).BalloonBackColor, AzkarFile
                
                WriteINI Section, "PlaySound", Azkars(x).PlaySound, AzkarFile
                WriteINI Section, "Animation", Azkars(x).Animation, AzkarFile
                WriteINI Section, "AnimationTime", Azkars(x).AnimationTime, AzkarFile
                WriteINI Section, "ShowDelay", Azkars(x).ShowDelay, AzkarFile
                WriteINI Section, "Transparency", Azkars(x).Transparency, AzkarFile
                WriteINI Section, "Theme", Azkars(x).Theme, AzkarFile
                WriteINI Section, "Alignment", Azkars(x).Alignment, AzkarFile
                WriteINI Section, "SoundFile", Azkars(x).SoundFile, AzkarFile
                
                WriteINI Section, "FontName", Azkars(x).FontName, AzkarFile
                WriteINI Section, "FontSize", Azkars(x).FontSize, AzkarFile
                WriteINI Section, "FontColor", Azkars(x).FontColor, AzkarFile
                
                WriteINI Section, "FontBold", Azkars(x).FontBold, AzkarFile
                WriteINI Section, "FontItalic", Azkars(x).FontItalic, AzkarFile
                WriteINI Section, "FontUnderline", Azkars(x).FontUnderline, AzkarFile
                
                Counter = Counter + 1
            End If
        Next
        
        LoadAzkarFile
        
        If iIndex >= lstAzkar.ListCount Then
            iIndex = lstAzkar.ListCount - 1
        End If
        
        lstAzkar.ListIndex = iIndex
        EditZekr iIndex
        
        StartAzkars
        
    End If

End Sub

Private Sub cmdBrowseForFile_Click()
    
    On Error GoTo ErrHandler
    
    With cdlDialog
        .FileName = ""
        .InitDir = AppPath + "Sound"
        .DialogTitle = "Open Sound File"
        .CancelError = False
        .Filter = "All Sound Types (*.wav;*.mp3)|*.wav;*.mp3|" + _
                    "All Files (*.*)|*.*"
        
        .ShowOpen
        
        If Len(.FileName) = 0 Then Exit Sub
        txtSoundFile = .FileName
    End With
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbOKOnly Or vbCritical, "Error"
End Sub

Private Sub cmdPlaySound_Click()
    
    Dim FileName As String
    
    FileName = Trim(txtSoundFile.Text)
    If FileExists(FileName) = False Then
        Exit Sub
    End If
    
    StopPlaying = False
    cmdStopSound.Enabled = True
    cmdPlaySound.Enabled = False
    
    MMC.Notify = False
    MMC.Wait = False
    MMC.Shareable = False
    MMC.DeviceType = "WaveAudio"
    MMC.FileName = FileName
    MMC.Command = "Open"
    MMC.Command = "Play"
    Do
        If MMC.Mode <> mciModePlay Then Exit Do
        DoEvents
    Loop While StopPlaying = False
    
    MMC.Command = "Stop"
    MMC.Command = "Close"
     
    cmdStopSound.Enabled = False
    cmdPlaySound.Enabled = True
End Sub

Private Sub cmdRecordSound_Click()
    
    cmdStopSound.Enabled = True
    cmdRecordSound.Enabled = False
    
    Dim FileName As String
    
    FileName = Trim(txtSoundFile.Text)
    If FileName = "" Then
        FileName = AppPath + "Sound" + "\Azkar_" + CStr(lstAzkar.ListIndex) + ".wav"
    End If
    
    txtSoundFile.Text = FileName
    
    MMC.Notify = False
    MMC.Wait = False
    MMC.Shareable = False
    MMC.DeviceType = "WaveAudio"
    MMC.FileName = FileName
    MMC.To = 10000
    MMC.Command = "Open"
    MMC.RecordMode = mciRecordOverwrite
    MMC.Command = "Record"

End Sub

Private Sub cmdRemoveSoundFile_Click()
    txtSoundFile.Text = ""
End Sub

Private Sub cmdStopSound_Click()
    StopPlaying = True
    If MMC.Mode = mciModePlay Then
        MMC.Command = "Stop"
        MMC.Command = "Close"
    ElseIf MMC.Mode = mciModeRecord Then
        MMC.Command = "Save"
        MMC.Command = "Close"
    End If
    cmdRecordSound.Enabled = True
    cmdStopSound.Enabled = False
End Sub

'====================================================================
'====================================================================
Sub PrepareAnimationList(Optional lIndex As Long = 0)

    cboAnimation.AddItem "Slide From Left"
    cboAnimation.AddItem "Slide From Right"
    cboAnimation.AddItem "Slide From Top"
    cboAnimation.AddItem "Slide From Bottom"
    cboAnimation.AddItem "Unfold Left Top"
    cboAnimation.AddItem "Unfold Left Bottom"
    cboAnimation.AddItem "Unfold Right Top"
    cboAnimation.AddItem "Unfold Right Bottom"
    cboAnimation.AddItem "Strech Horizontally"
    cboAnimation.AddItem "Strech Vertically"
    cboAnimation.AddItem "Zoom Out"
    cboAnimation.AddItem "Fold Out"
    cboAnimation.AddItem "Curton Horizontal"
    cboAnimation.AddItem "Curton Vertical"
    
End Sub
    
Sub PreparePopupSize(Optional lIndex As Long = 0)
    
    cboPopupWidth.Clear
    cboPopupWidth.AddItem "25"
    cboPopupWidth.AddItem "50"
    cboPopupWidth.AddItem "75"
    cboPopupWidth.AddItem "100"
    cboPopupWidth.AddItem "125"
    cboPopupWidth.AddItem "150"
    cboPopupWidth.AddItem "175"
    cboPopupWidth.AddItem "200"
    cboPopupWidth.AddItem "225"
    cboPopupWidth.AddItem "250"
    cboPopupWidth.AddItem "275"
    cboPopupWidth.AddItem "300"
    cboPopupWidth.AddItem "325"
    cboPopupWidth.AddItem "350"
    cboPopupWidth.AddItem "375"
    cboPopupWidth.AddItem "400"
    
    cboPopupWidth.ListIndex = lIndex
    
    cboPopupHeight.Clear
    cboPopupHeight.AddItem "25"
    cboPopupHeight.AddItem "50"
    cboPopupHeight.AddItem "75"
    cboPopupHeight.AddItem "100"
    cboPopupHeight.AddItem "125"
    cboPopupHeight.AddItem "150"
    cboPopupHeight.AddItem "175"
    cboPopupHeight.AddItem "200"
    cboPopupHeight.AddItem "225"
    cboPopupHeight.AddItem "250"
    cboPopupHeight.AddItem "275"
    cboPopupHeight.AddItem "300"
    cboPopupHeight.AddItem "325"
    cboPopupHeight.AddItem "350"
    cboPopupHeight.AddItem "375"
    cboPopupHeight.AddItem "400"
    cboPopupHeight.ListIndex = lIndex
    
End Sub

Private Sub PrepareFiles()
    
    On Error GoTo ErrHandler
    
    'DirExists AppPath + "data"
    
    Dim FileNum As Long
    If FileExists(AzkarFile) = False Then
        FileNum = FreeFile
        Open AzkarFile For Output As #FileNum
        Close #FileNum
    End If
    
    Exit Sub
    
ErrHandler:
    Dim ErrRet As Long
    ErrRet = MsgBox("Error: " & Err.Description, vbRetryCancel Or vbCritical, "Error")
    If ErrRet = vbRetry Then
        Resume Next
    End If
End Sub

'====================================================================
Function LoadBitmap(Path As String) As Long
    LoadBitmap = LoadImage(App.hInstance, App.Path + "\" + Path, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
End Function

Function LoadIcon(Path As String, cx As Long, cy As Long) As Long
    LoadIcon = LoadImage(App.hInstance, App.Path + "\" + Path, IMAGE_ICON, cx, cy, LR_LOADFROMFILE)
End Function

Private Sub ShowBalloon(ByVal enIconType As blIconType, ByVal sPrompt As String, Optional ByVal sTitle As String, Optional ByVal lTimeout As Long, _
                        Optional ByVal lBackColor As Long = -1, Optional ByVal lForeColor As Long = -1)
  On Error GoTo lblErr
  Dim lX As Long, lY As Long
    
  Call TrayIcon.GetIconMiddle(lX, lY)
  If lForeColor = -1 Then lForeColor = vbBlack
  If lBackColor = -1 Then lBackColor = &H80000018
  
  With ToolTipOnDemand
    .ParentHwnd = TrayIcon.SysTrayHWnd
    .x = lX
    .y = lY
    .BackColor = lBackColor
    .ForeColor = lForeColor
    .Prompt = sPrompt
    .Title = sTitle
    .TimeOut = lTimeout
    .IconType = enIconType
    .Show
  End With

lblExit:
  Exit Sub
  
lblErr:
  'MsgBox "Error #" & Err.Number & " in " & Err.Source & vbCrLf & Err.Description, vbCritical
  Resume lblExit
End Sub

Private Sub lstAzkar_Click()
    If bLoaded = False Then Exit Sub
    EditZekr Azkars(lstAzkar.ListIndex).ID
End Sub

Private Sub picPosition_Click(index As Integer)
    Dim x As Long
    
    For x = 0 To 8
        picPosition(x).BackColor = &H8000000F
    Next
    
    picPosition(index).BackColor = vbBlue
    
End Sub

Private Sub Tray_LButtonDblClick()
    ZoomFrm Me.hwnd, BottomRightZoom, False
    Me.Show
End Sub

Private Sub Tray_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Debug.Print "Tray_MouseMove: "; Button, Shift, X, Y
End Sub

Private Sub Tray_RButtonDown()

    'Display popup menu when user presses the right mouse button on
    'the System Tray icon
    PopupMenu Me.mnuTrayPopup
    
End Sub

Private Sub mnuTrayShow_Click()
    ZoomFrm Me.hwnd, BottomRightZoom, False
    Me.Show
End Sub

Private Sub mnuTrayHide_Click()
    ZoomFrm Me.hwnd, BottomRightZoom, True
    Me.Hide
End Sub

Private Sub mnuTrayExit_Click()
    Unload Me
End Sub

Private Sub cboFontName_Click()
    
    If bLoaded = False Then Exit Sub
    
    Dim FontName As String
    
    FontName = cboFontName.List(cboFontName.ListIndex)
    txtZekr.SelStart = 0
    txtZekr.SelLength = Len(txtZekr.Text)
    txtZekr.Font.Name = FontName
    'txtZekr.SelColor = cmdColor.BackColor
    txtZekr.SelStart = 0
    txtZekr.SelLength = 0
    
End Sub

Private Sub cboFontSize_Click()
    If bLoaded = False Then Exit Sub
    txtZekr.SelStart = 0
    txtZekr.SelLength = Len(txtZekr.Text)
    txtZekr.Font.Size = cboFontSize.List(cboFontSize.ListIndex)
    'txtZekr.SelColor = cmdColor.BackColor
    txtZekr.SelStart = 0
    txtZekr.SelLength = 0
End Sub

Private Sub cmdPaste_Click()
    If bLoaded = False Then Exit Sub
    txtZekr.SelText = Clipboard.GetText()
    
    txtZekr.SelStart = 0
    txtZekr.SelLength = Len(txtZekr.Text)
    'txtZekr.SelColor = cmdColor.BackColor
    txtZekr.SelStart = 0
    txtZekr.SelLength = 0
  
    'txtZekr.SetFocus
End Sub

Private Sub cmdCopy_Click()
    If bLoaded = False Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText txtZekr.SelText
End Sub

Private Sub cmdCut_Click()
    If bLoaded = False Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText txtZekr.SelText
    txtZekr.SelText = ""
    'txtZekr.SetFocus
    txtZekr.SelStart = 0
    txtZekr.SelLength = 0
End Sub

Private Sub cmdFontBold_Click()
    If bLoaded = False Then Exit Sub
    txtZekr.SelStart = 0
    txtZekr.SelLength = Len(txtZekr.Text)
    txtZekr.Font.Bold = Not txtZekr.Font.Bold
    'txtZekr.SelColor = cmdColor.BackColor
    'txtZekr.SetFocus
    txtZekr.SelStart = 0
    txtZekr.SelLength = 0
End Sub

Private Sub cmdFontItalic_Click()
    If bLoaded = False Then Exit Sub
    txtZekr.SelStart = 0
    txtZekr.SelLength = Len(txtZekr.Text)
    txtZekr.Font.Italic = Not txtZekr.Font.Italic
    'txtZekr.SelColor = cmdColor.BackColor
    'txtZekr.SetFocus
    txtZekr.SelStart = 0
    txtZekr.SelLength = 0
End Sub

Private Sub cmdFontUnderline_Click()
    If bLoaded = False Then Exit Sub
    txtZekr.SelStart = 0
    txtZekr.SelLength = Len(txtZekr.Text)
    txtZekr.Font.Underline = Not txtZekr.Font.Underline
    'txtZekr.SelColor = cmdColor.BackColor
    'txtZekr.SetFocus
    txtZekr.SelStart = 0
    txtZekr.SelLength = 0
End Sub

Private Sub picBalloonColor_Click()
    'cdlDialog.Flags = cdlCFScreenFonts
    cdlDialog.CancelError = True
    On Error Resume Next
    
    With cdlDialog
        '.Color = WatermarkStyleColor
    End With
    
    cdlDialog.ShowColor
    If Err.Number = cdlCancel Then Exit Sub
    
    ' Use the dialog's properties.
    With cdlDialog
        picBalloonColor.BackColor = .Color
    End With
End Sub

Private Sub picBalloonBackColor_Click()
    'cdlDialog.Flags = cdlCFScreenFonts
    cdlDialog.CancelError = True
    On Error Resume Next
    
    With cdlDialog
        '.Color = WatermarkStyleColor
    End With
    
    cdlDialog.ShowColor
    If Err.Number = cdlCancel Then Exit Sub
    
    ' Use the dialog's properties.
    With cdlDialog
        picBalloonBackColor.BackColor = .Color
    End With
End Sub

Private Sub picTheme_Click()
    
    If bLoaded = False Then Exit Sub
    
    On Error GoTo ErrHandler
    
    frmTheme.SelectedImage = picTheme.Tag
    
    frmTheme.Show vbModal, Me
    
    If NewSelectedImage <> "" Then
        picTheme.Tag = NewSelectedImage
        cmdSaveZekr_Click
        Set picTheme.Picture = LoadPicture(AppPath + "themes\" + NewSelectedImage)
        picTheme.Refresh
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " + Err.Description, vbCritical
End Sub

Private Sub chkPlaySound_Click()

    If chkPlaySound.Value Then
        txtSoundFile.Enabled = True
        txtSoundFile.Refresh
        cmdBrowseForFile.Enabled = True
        cmdRemoveSoundFile.Enabled = True
        cmdRecordSound.Enabled = True
        'cmdStopSound.Enabled = True
        cmdPlaySound.Enabled = True
    Else
        txtSoundFile.Enabled = False
        cmdBrowseForFile.Enabled = False
        cmdRemoveSoundFile.Enabled = False
        cmdRecordSound.Enabled = False
        'cmdStopSound.Enabled = False
        cmdPlaySound.Enabled = False
        txtSoundFile.Refresh
    End If

End Sub

Private Sub cmdFont_Click()
    
    cdlDialog.flags = cdlCFScreenFonts
    cdlDialog.CancelError = True
    On Error Resume Next
    
    With cdlDialog
        '.FontName = WatermarkFontname
        '.FontSize = WatermarkFontSize
        '.FontBold = WatermarkFontBold
        '.FontItalic = WatermarkFontItalic
    End With
    
    cdlDialog.ShowFont
    If Err.Number = cdlCancel Then Exit Sub
    
    ' Use the dialog's properties.
    With cdlDialog
        'WatermarkFontname = .FontName
        'WatermarkFontSize = .FontSize
        'WatermarkFontBold = .FontBold
        'WatermarkFontItalic = .FontItalic
    End With
    
End Sub

Private Sub cmdColor_Click()
    
    'cdlDialog.Flags = cdlCFScreenFonts
    cdlDialog.CancelError = True
    On Error Resume Next
    
    With cdlDialog
        '.Color = WatermarkStyleColor
    End With
    
    cdlDialog.ShowColor
    If Err.Number = cdlCancel Then Exit Sub
    
    ' Use the dialog's properties.
    With cdlDialog
        txtZekr.SelStart = 0
        txtZekr.SelLength = Len(txtZekr.Text)
        'txtZekr.SelColor = .Color
        cmdColor.BackColor = .Color
    End With
    
End Sub

Public Sub StopAzkars()
    
    On Error Resume Next
    
    Dim uct As Object
        
    For Each uct In ucTimer
        uct.Enabled = False
        If uct.index > 0 Then
            Unload uct
        End If
    Next
    
    ucTimerRotateAzkar.Enabled = False
    
    RunningStatus = False
    
    mnuTryDisable.Caption = MenuTrayEnable
    
End Sub

Public Sub StartAzkars()
    
    Dim x As Long
    Dim Counter As Long
    
    If AzkarCount = 0 Then Exit Sub
    
    If RotateAzkar = 0 Then
        Counter = 0
        For x = LBound(Azkars) To UBound(Azkars)
            If Azkars(x).Enabled And Azkars(x).Zekr <> "" Then
                Counter = Counter + 1
                Load ucTimer(Counter)
                ucTimer(Counter).Interval = Azkars(x).Period * 60
                ucTimer(Counter).Enabled = False
                ucTimer(Counter).Enabled = True
            End If
        Next
    Else
        ucTimerRotateAzkar.Interval = RotateAzkarPeriod * 60
        ucTimerRotateAzkar.Enabled = True
    End If
    
    mnuTryDisable.Caption = MenuTrayDisable
    
    RunningStatus = True
    
End Sub

Private Sub TrayIcon_TrayIconMoved(ByVal lX As Long, ByVal lY As Long)
    ToolTipOnDemand.x = lX
    ToolTipOnDemand.y = lY
End Sub

Private Sub ToolTipOnDemand_BalloonDestroyed()
    TrayIcon.TrackIconMovement = False
End Sub

Private Sub ToolTipOnDemand_BalloonShowed()
    TrayIcon.TrackIconMovement = True
End Sub

Private Sub TrayIcon_BalloonClick(ByVal MouseEvent As stBalloonClickType)
  
    Select Case MouseEvent
        Case stbBalloonShow
        Case stbBalloonHide
        Case stbRightClick
        Case stbLeftClick
    End Select

End Sub

Private Sub ToolTipOnDemand_MouseEvents(MouseEvent As Long)
  
  Select Case MouseEvent
  
    Case stMouseMove
    
    Case stLeftButtonDown
      ToolTipOnDemand.Destroy
      
    Case stLeftButtonUp
      
    Case stLeftButtonDoubleClick
      
    Case stRightButtonDown
      ToolTipOnDemand.Destroy
      
    Case stRightButtonUp
    Case stRightButtonDoubleClick
    Case stMiddleButtonDown
    Case stMiddleButtonUp
    Case stMiddleButtonDoubleClick
  End Select

End Sub

Private Sub ucTimer_Timer(index As Integer)
    
    'Debug.Print "ucTimer: " + CStr(Index)
    If index < 1 Then Exit Sub
    ShowZekr CLng(index - 1)
    
End Sub

Sub ShowZekr(index As Long)

    Dim Counter As Long
    Dim Popups As frmPopup
    Dim x As Long, y As Long
    Dim lWidth As Long, lHeight As Long
    
    Counter = index
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    If Azkars(Counter).Enabled = 0 Then Exit Sub
    '----------------------------------------------------------------
    '   Show Tray Icon Balloon
    '----------------------------------------------------------------
    If Azkars(Counter).ShowBalloon Then
        ShowBalloon blNoIcon, Azkars(Counter).Zekr, "", Azkars(Counter).ShowDelay * 1000, Azkars(Counter).BalloonBackColor, Azkars(Counter).BalloonColor
    End If
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    '   Show Popup Window
    '----------------------------------------------------------------
    If Azkars(Counter).ShowPopup Then
        Set Popups = New frmPopup
        MakeTopMostNoFocus Popups.hwnd
        
        'ScreenWidth = GetSystemMetrics(SM_CXFULLSCREEN)
        'ScreenHeight = GetSystemMetrics(SM_CYFULLSCREEN)
        
        lWidth = (Azkars(Counter).PopupWidth + 1) * 25 * Screen.TwipsPerPixelX
        lHeight = (Azkars(Counter).PopupHeight + 1) * 25 * Screen.TwipsPerPixelY
        
        Select Case Azkars(Counter).PopupPosition
    
            Case 0:
                'cboPopupPosition.AddItem "Top Left"
                x = 0
                y = 0
            Case 1:
                'cboPopupPosition.AddItem "Top Center"
                x = ScreenWidth / 2 - lWidth / 2
                y = 0
            Case 2:
                'cboPopupPosition.AddItem "Top Right"
                x = ScreenWidth - lWidth
                y = 0
                
            Case 7:
                'cboPopupPosition.AddItem "Bottom Center"
                x = ScreenWidth / 2 - lWidth / 2
                y = ScreenHeight - lHeight
    
            Case 6:
                x = 0
                'cboPopupPosition.AddItem "Bottom Left"
                y = ScreenHeight - lHeight
    
            Case 8:
                'cboPopupPosition.AddItem "Bottom Right"
                x = ScreenWidth - lWidth
                y = ScreenHeight - lHeight
    
            Case 4:
                '    cboPopupPosition.AddItem "Center"
                x = ScreenWidth / 2 - lWidth / 2
                y = ScreenHeight / 2 - lHeight / 2
    
            Case 3:
                x = 0
                '    cboPopupPosition.AddItem "Center Left"
                y = ScreenHeight / 2 - lHeight / 2
    
            Case 5:
                '    cboPopupPosition.AddItem "Center Right"
                x = ScreenWidth - lWidth
                y = ScreenHeight / 2 - lHeight / 2
    
                
        End Select
         
        Popups.Move x, y, lWidth, lHeight
        
        Popups.lblMessage.Move 200, 200, Popups.Width - 400, Popups.Height - 400
        
        Popups.lblMessage.Font.Name = Azkars(Counter).FontName
        Popups.lblMessage.Font.Size = Azkars(Counter).FontSize
        Popups.lblMessage.ForeColor = Azkars(Counter).FontColor
        Popups.lblMessage.Font.Bold = Azkars(Counter).FontBold
        Popups.lblMessage.Font.Italic = Azkars(Counter).FontItalic
        Popups.lblMessage.Font.Underline = Azkars(Counter).FontUnderline
        Popups.lblMessage.Alignment = Azkars(Counter).Alignment
                
        Popups.lblMessage.Caption = Azkars(Counter).Zekr
        Popups.PopupShape.Move Popups.PopupShape.BorderWidth, Popups.PopupShape.BorderWidth, Popups.Width - (2 * Popups.PopupShape.BorderWidth), Popups.Height - (2 * Popups.PopupShape.BorderWidth)
        
        'Tray.TipText = Me.Caption + " - " + Azkars(Counter).Zekr
        If Azkars(Counter).Zekr <> "" Then
            TrayIcon.ToolTip = Azkars(Counter).Zekr
        Else
            TrayIcon.ToolTip = Me.Caption
        End If
        
        If Azkars(Counter).Theme <> "" Then
            If FileExists(AppPath + "themes\" + Azkars(Counter).Theme) = True Then
                Set Popups.Picture = LoadPicture(AppPath + "themes\" + Azkars(Counter).Theme)
            End If
        End If
        
        'Debug.Print X, Y, lWidth, lHeight
                    
        ShowWindow Popups.hwnd, SW_SHOWNOACTIVATE
       
        AnimateForm Popups, aload, Azkars(Counter).Animation, Azkars(Counter).AnimationTime / 30, 30
        
        Popups.Effect = Azkars(Counter).Animation
        Popups.FrameTime = Azkars(Counter).AnimationTime / 30
        Popups.FrameCount = 30
        
        FadeForm Popups.hwnd, Azkars(Counter).Transparency
        
        Popups.tmrClose.Interval = Azkars(Counter).ShowDelay * 1000
        Popups.tmrClose.Enabled = True
        Set Popups = Nothing
    End If
    '----------------------------------------------------------------
    '   Play sound
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    If Azkars(Counter).SoundFile <> "" And Azkars(Counter).PlaySound Then
        MMC.Command = "Stop"
        MMC.Command = "Close"
        MMC.Notify = False
        MMC.Wait = False
        MMC.Shareable = False
        MMC.DeviceType = "WaveAudio"
        MMC.FileName = Azkars(Counter).SoundFile
        MMC.Command = "Open"
        MMC.Command = "Play"
    End If
    '----------------------------------------------------------------

End Sub

Private Sub MMC_Done(NotifyCode As Integer)
    'Debug.Print NotifyCode
End Sub

Private Sub cmdShowZekr_Click()
    Dim lEnabled As Long
    
    lEnabled = chkEnabled.Value
    
    If lEnabled = 0 Then
        chkEnabled.Value = 1
    End If
    
    cmdSaveZekr_Click
    
    ShowZekr lstAzkar.ListIndex
End Sub

Sub ShowRandomZekr()
    ShowZekr CLng((AzkarCount - 1) * Rnd)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'UnloadMode possibilities:
'0   The user has chosen the Close command from the Control-menu box on the form.
'1   The Unload method has been invoked from code.
'2   The current Windows-environment session is ending.
'3   The Microsoft Windows Task Manager is closing the application.
'4   An MDI child form is closing because the MDI form is closing.
    
'    If UnloadMode = vbFormControlMenu Then
'        ZoomFrm Me.hWnd, BottomRightZoom, True
'        Tray.ShowIcon
'        Me.Hide
'        Cancel = 1
'    End If

    If UnloadMode = vbFormControlMenu Then
        If TrayIcon.Created Then
            Cancel = 1
            HideForm
        End If
    End If

End Sub

Private Sub HideForm()
  Me.Hide
  ZoomFrm Me.hwnd, BottomRightZoom, True
  ShowBalloon blNoIcon, Me.Caption, "", 3000, vbRed, vbYellow
End Sub

Sub EndProgram()

    'Tray.DeleteIcon
    'Set Tray = Nothing
    On Error GoTo ErrHandler
    
    StopAzkars
    
    If TrayIcon.Created Then
        TrayIcon.Remove
        ToolTipOnDemand.Destroy
    End If
    
    Dim i As Integer
    
    'close all sub forms
    For i = Forms.count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    
    If Me.WindowState <> vbMinimized Then
        SaveSettings AppRegPath, "Settings", "MainLeft", Me.Left
        SaveSettings AppRegPath, "Settings", "MainTop", Me.Top
        SaveSettings AppRegPath, "Settings", "MainWidth", Me.Width
        SaveSettings AppRegPath, "Settings", "MainHeight", Me.Height
    End If
    
    Exit Sub
ErrHandler:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndProgram
End Sub

Private Sub cmdClose_Click()
    HideForm
End Sub

Private Sub cmdExit_Click()
    Dim ret As Long
    ret = MsgBox(ExitMessage, vbExclamation Or vbYesNoCancel)
    If ret = vbYes Then
        Unload Me
    End If
End Sub

Private Sub TrayIcon_TrayMouseEvent(ByVal MouseEvent As stMouseEvent)
  
  Select Case MouseEvent
    Case stMouseMove
        'ShowRandomZekr
    
    Case stLeftButtonDown
        'ShowRandomZekr
        
    Case stLeftButtonUp
        'ShowRandomZekr
        
    Case stLeftButtonDoubleClick
        ZoomFrm Me.hwnd, BottomRightZoom, False
        Me.Show
    
    Case stRightButtonDown
      PopupMenu mnuTrayPopup
      
    Case stRightButtonUp
    Case stRightButtonDoubleClick
    
    Case stMiddleButtonDown
        ShowRandomZekr
        
    Case stMiddleButtonUp
    Case stMiddleButtonDoubleClick
  End Select

End Sub

Private Sub tmrHide_Timer()
    
    tmrHide.Enabled = False
    'Tray.ShowIcon
    'ZoomFrm Me.hwnd, BottomRightZoom, True
    HideForm
    
End Sub

Private Sub ucTimerRotateAzkar_Timer()
    
    RotateAzkarIndex = RotateAzkarIndex + 1
    If RotateAzkarIndex > (AzkarCount - 1) Then
        RotateAzkarIndex = 0
    End If
    
    Do While Azkars(RotateAzkarIndex).Enabled = 0
        If RotateAzkarIndex > (AzkarCount - 1) Then
            RotateAzkarIndex = 0
            Exit Do
        End If
        RotateAzkarIndex = RotateAzkarIndex + 1
    Loop
    
    If Azkars(RotateAzkarIndex).Enabled = 0 Then
        RotateAzkarIndex = 0
        Exit Sub
    End If
    
    ShowZekr RotateAzkarIndex
    
End Sub
