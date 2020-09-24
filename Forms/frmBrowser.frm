VERSION 5.00
Object = "{C151518A-D64D-4C66-96F3-DB69BF286B30}#1.0#0"; "WinXPC Engine.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl MMC 
      Height          =   330
      Left            =   945
      TabIndex        =   12
      Top             =   4590
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   900
      Top             =   3825
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.PictureBox picBottom 
      Height          =   915
      Left            =   180
      ScaleHeight     =   855
      ScaleWidth      =   7965
      TabIndex        =   1
      Top             =   5490
      Width           =   8025
      Begin VB.CommandButton cmdCounterUp100 
         Caption         =   "+"
         Height          =   330
         Left            =   5760
         TabIndex        =   11
         Top             =   135
         Width           =   645
      End
      Begin VB.CommandButton cmdCounterReset100 
         Caption         =   "00"
         Height          =   330
         Left            =   4140
         TabIndex        =   9
         Top             =   135
         Width           =   645
      End
      Begin VB.CommandButton cmdCounterUp 
         Caption         =   "+"
         Height          =   330
         Left            =   2340
         TabIndex        =   7
         Top             =   135
         Width           =   645
      End
      Begin VB.CommandButton cmdResetCounter 
         Caption         =   "00"
         Height          =   330
         Left            =   720
         TabIndex        =   5
         Top             =   135
         Width           =   645
      End
      Begin VB.CommandButton cmdAzkarIndex 
         Caption         =   "Index"
         Height          =   285
         Left            =   5760
         TabIndex        =   3
         Top             =   585
         Width           =   1050
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   285
         Left            =   6885
         TabIndex        =   2
         Top             =   585
         Width           =   1050
      End
      Begin VB.Label lblCounter100 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   330
         Left            =   4770
         TabIndex        =   10
         Top             =   135
         Width           =   1005
      End
      Begin VB.Label lblCounter100Label 
         Caption         =   "100 Counter:"
         Height          =   240
         Left            =   3195
         TabIndex        =   8
         Top             =   180
         Width           =   915
      End
      Begin VB.Label lblCounter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "33"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   1350
         TabIndex        =   6
         Top             =   135
         Width           =   1005
      End
      Begin VB.Label lblCounterLabel 
         Caption         =   "Counter:"
         Height          =   240
         Left            =   45
         TabIndex        =   4
         Top             =   180
         Width           =   645
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   3300
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4785
      ExtentX         =   8440
      ExtentY         =   5821
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HelpFile As String
Dim BrowserZoom As Long
Dim bLoaded As Boolean
Dim Counter As Long
Dim Counter100 As Long

Private Sub Form_Load()

    On Error Resume Next
    bLoaded = False
    '----------------------------------------------------------------
    Me.Left = GetSettings(AppRegPath, "Settings", "BrowserLeft", 1000)
    Me.Top = GetSettings(AppRegPath, "Settings", "BrowserTop", 1000)
    Me.Width = GetSettings(AppRegPath, "Settings", "BrowserWidth", 9000)
    Me.Height = GetSettings(AppRegPath, "Settings", "BrowserHeight", 7500)
    '----------------------------------------------------------------
    lblCounterLabel.Caption = CounterLabel
    cmdResetCounter.Caption = ResetCounter
    cmdCounterUp.Caption = CounterUp
    lblCounter100Label.Caption = Counter100Label
    cmdCounterReset100.Caption = CounterReset100
    cmdCounterUp100.Caption = CounterUp100
    cmdAzkarIndex.Caption = AzkarIndex
    cmdClose.Caption = AzkarClose
    '----------------------------------------------------------------
    HelpFile = AppPath & "Azkar\Azkar.html"
    If HelpFile <> "" Then
        WebBrowser.Navigate2 HelpFile
    Else
        WebBrowser.Navigate2 "about:blank"
    End If
    '----------------------------------------------------------------
    Counter = 0
    lblCounter.Caption = Counter
    Counter100 = 0
    lblCounter100.Caption = Counter100
    '----------------------------------------------------------------
    Me.Caption = fMainForm.Caption
    Me.RightToLeft = fMainForm.RightToLeft
    '----------------------------------------------------------------
    WindowsXPC1.ColorScheme = XP_OliveGreen '= System ' = XP_Blue
    WindowsXPC1.InitSubClassing
    '----------------------------------------------------------------
    WebBrowser.SetFocus
    bLoaded = True
End Sub

Private Sub cmdCounterReset100_Click()
    Counter100 = 0
    lblCounter100.Caption = Counter100
End Sub

Private Sub cmdCounterUp_Click()
    Counter = Counter + 1
    lblCounter.Caption = Counter
End Sub

Private Sub cmdCounterUp100_Click()
    Counter100 = Counter100 + 1
    
    If Counter100 = 33 Then
        MMC.Command = "Stop"
        MMC.Command = "Close"
        MMC.Notify = False
        MMC.Wait = False
        MMC.Shareable = False
        MMC.DeviceType = "WaveAudio"
        MMC.FileName = AppPath + "Sound\tick33.wav"
        MMC.Command = "Open"
        MMC.Command = "Play"
    ElseIf Counter100 = 66 Then
        MMC.Command = "Stop"
        MMC.Command = "Close"
        MMC.Notify = False
        MMC.Wait = False
        MMC.Shareable = False
        MMC.DeviceType = "WaveAudio"
        MMC.FileName = AppPath + "Sound\tick66.wav"
        MMC.Command = "Open"
        MMC.Command = "Play"
    ElseIf Counter100 >= 100 Then
        Counter100 = 0
        MMC.Command = "Stop"
        MMC.Command = "Close"
        MMC.Notify = False
        MMC.Wait = False
        MMC.Shareable = False
        MMC.DeviceType = "WaveAudio"
        MMC.FileName = AppPath + "Sound\tick100.wav"
        MMC.Command = "Open"
        MMC.Command = "Play"
    End If
    
    lblCounter100.Caption = Counter100
End Sub

Private Sub cmdResetCounter_Click()
    Counter = 0
    lblCounter.Caption = Counter
End Sub

Private Sub cmdAzkarIndex_Click()
    If HelpFile <> "" Then
        WebBrowser.Navigate2 HelpFile
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    
    If Me.WindowState <> vbMinimized Then
        picBottom.Move 0, Me.ScaleHeight - picBottom.Height, Me.ScaleWidth - 0
        
        cmdClose.Move picBottom.ScaleWidth - cmdClose.Width
        cmdAzkarIndex.Move picBottom.ScaleWidth - cmdAzkarIndex.Width - cmdClose.Width - 10
        
        WebBrowser.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - picBottom.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Me.WindowState <> vbMinimized Then
        SaveSettings AppRegPath, "Settings", "BrowserLeft", Me.Left
        SaveSettings AppRegPath, "Settings", "BrowserTop", Me.Top
        SaveSettings AppRegPath, "Settings", "BrowserWidth", Me.Width
        SaveSettings AppRegPath, "Settings", "BrowserHeight", Me.Height
    End If

End Sub
