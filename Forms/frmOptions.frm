VERSION 5.00
Object = "{C151518A-D64D-4C66-96F3-DB69BF286B30}#1.0#0"; "WinXPC Engine.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOptions.frx":0000
   ScaleHeight     =   3345
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   5400
      Top             =   2790
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picLabelsColor 
      Height          =   375
      Left            =   3330
      ScaleHeight     =   315
      ScaleWidth      =   540
      TabIndex        =   12
      Top             =   1935
      Width           =   600
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   1395
      Top             =   3060
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   3105
      TabIndex        =   10
      Top             =   2610
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   285
      Left            =   1755
      TabIndex        =   9
      Top             =   2610
      Width           =   1185
   End
   Begin VB.TextBox txtRotateAzkarPeriod 
      Height          =   285
      Left            =   3330
      TabIndex        =   6
      Top             =   1440
      Width           =   1185
   End
   Begin VB.CheckBox chkRotateAzkar 
      Caption         =   "Check1"
      Height          =   190
      Left            =   3330
      TabIndex        =   4
      Top             =   1125
      Width           =   190
   End
   Begin VB.CheckBox chkEnabled 
      Height          =   190
      Left            =   3330
      TabIndex        =   2
      Top             =   405
      Width           =   190
   End
   Begin VB.CheckBox chkAutoStartUp 
      Height          =   190
      Left            =   3330
      TabIndex        =   0
      Top             =   720
      Width           =   190
   End
   Begin VB.Label lblLabelsColor 
      BackStyle       =   0  'Transparent
      Caption         =   "Labels color:"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   315
      TabIndex        =   11
      Top             =   1935
      Width           =   2220
   End
   Begin VB.Label lblRotationTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Rotation Time:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   990
      TabIndex        =   8
      Top             =   1440
      Width           =   2085
   End
   Begin VB.Label lblMinutes 
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4545
      TabIndex        =   7
      Top             =   1440
      Width           =   1635
   End
   Begin VB.Label lblRotateAzkar 
      BackStyle       =   0  'Transparent
      Caption         =   "Rotate all Azkar:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   315
      TabIndex        =   5
      Top             =   1125
      Width           =   2940
   End
   Begin VB.Label lblAzkarStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Enabled:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   450
      Width           =   2535
   End
   Begin VB.Label lblAutoStart 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto start with Windows:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   315
      TabIndex        =   1
      Top             =   765
      Width           =   2940
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bLoaded As Boolean

Private Sub Form_Load()
    
    bLoaded = False
    
    Me.Icon = fMainForm.Icon
    Me.Caption = fMainForm.Caption
    
    Me.RightToLeft = fMainForm.RightToLeft
    
    If RunningStatus = True Then
        chkEnabled.Value = 1
    Else
        chkEnabled.Value = 0
    End If
    
    chkRotateAzkar.Value = RotateAzkar
    
    If RotateAzkar = 1 Then
        txtRotateAzkarPeriod.Enabled = True
        lblRotationTime.Enabled = True
        lblMinutes.Enabled = True
    Else
        txtRotateAzkarPeriod.Enabled = False
        lblRotationTime.Enabled = False
        lblMinutes.Enabled = False
    End If
    
    If RotateAzkarPeriod <= 0 Then RotateAzkarPeriod = 2
    txtRotateAzkarPeriod.Text = CStr(RotateAzkarPeriod)
    
    lblAzkarStatus.Caption = AzkarStatusLabel
    lblAutoStart.Caption = AutoStartMessage
    lblMinutes.Caption = MinutesLabel
    lblRotateAzkar.Caption = RotateAzkarLabel
    lblRotationTime.Caption = RotationTimeLabel
    lblLabelsColor.Caption = LabelsColorLabel
    cmdOK.Caption = CommandOK
    cmdCancel.Caption = CommandCancel
    
    chkAutoStartUp.Value = AutoStartUp
    '----------------------------------------------------------------
    picLabelsColor.BackColor = LabelsColor
    Dim Ctl  As Control
    For Each Ctl In Me.Controls
        If TypeOf Ctl Is Label Then
            Ctl.ForeColor = LabelsColor
        End If
    Next
    '----------------------------------------------------------------
    WindowsXPC1.ColorScheme = XP_OliveGreen '= System ' = XP_Blue
    WindowsXPC1.InitSubClassing
    '----------------------------------------------------------------
    bLoaded = True
    
End Sub

Private Sub chkRotateAzkar_Click()
    
    If chkRotateAzkar.Value = 1 Then
        txtRotateAzkarPeriod.Enabled = True
        lblRotationTime.Enabled = True
        lblMinutes.Enabled = True
    Else
        txtRotateAzkarPeriod.Enabled = False
        lblRotationTime.Enabled = False
        lblMinutes.Enabled = False
    End If
    
End Sub

Private Sub cmdOK_Click()

    '----------------------------------------------------------------
    RotateAzkar = chkRotateAzkar.Value
    RotateAzkarPeriod = Val(txtRotateAzkarPeriod.Text)
    If RotateAzkarPeriod <= 0 Then RotateAzkarPeriod = 2
    
    SaveSettings AppRegPath, "Settings", "RotateAzkar", RotateAzkar
    SaveSettings AppRegPath, "Settings", "RotateAzkarPeriod", RotateAzkarPeriod
    '----------------------------------------------------------------
    AutoStartUp = chkAutoStartUp.Value
    SaveSettings AppRegPath, "Settings", "AutoStartUp", AutoStartUp
    If AutoStartUp = 1 Then
        AutoRun = eAlways
    Else
        AutoRun = eNever
    End If
    '----------------------------------------------------------------
    SaveSettings AppRegPath, "Settings", "AutoStartUp", AutoStartUp
    
    fMainForm.StopAzkars
    If chkEnabled.Value = 1 Then
        fMainForm.StartAzkars
    Else
        fMainForm.StopAzkars
    End If
    '----------------------------------------------------------------
    LabelsColor = picLabelsColor.BackColor
    SaveSettings AppRegPath, "Settings", "LabelsColor", LabelsColor
    
    Dim Ctl  As Control
    
    For Each Ctl In fMainForm.Controls
        If TypeOf Ctl Is Label Then
            Ctl.ForeColor = LabelsColor
        End If
    Next
    
    For Each Ctl In Me.Controls
        If TypeOf Ctl Is Label Then
            Ctl.ForeColor = LabelsColor
        End If
    Next
    '----------------------------------------------------------------
    
    Unload Me
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub picLabelsColor_Click()
    
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
        picLabelsColor.BackColor = .Color
    End With
End Sub
