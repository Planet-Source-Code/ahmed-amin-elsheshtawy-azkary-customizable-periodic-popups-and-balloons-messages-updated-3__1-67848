VERSION 5.00
Object = "{C151518A-D64D-4C66-96F3-DB69BF286B30}#1.0#0"; "WinXPC Engine.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTheme 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmTheme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   6975
      Top             =   5175
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      PictureControl  =   0   'False
   End
   Begin VB.Timer tmrLoad 
      Interval        =   10
      Left            =   8910
      Top             =   2160
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   8685
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   285
      Left            =   8685
      TabIndex        =   2
      Top             =   315
      Width           =   1095
   End
   Begin VB.PictureBox picView 
      AutoSize        =   -1  'True
      Height          =   6000
      Left            =   2610
      ScaleHeight     =   5940
      ScaleWidth      =   5940
      TabIndex        =   1
      Top             =   90
      Width           =   6000
   End
   Begin MSComctlLib.ImageList imglImages 
      Left            =   7245
      Top             =   4275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwImages 
      Height          =   6045
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   10663
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedImage As String
Dim ImageFiles() As String

Private Sub Form_Load()
    
    Me.MousePointer = vbHourglass
    
    cmdOK.Caption = CommandOK
    cmdCancel.Caption = CommandCancel
    
    Me.Caption = fMainForm.Caption
    Me.RightToLeft = fMainForm.RightToLeft
    
    Me.MousePointer = vbDefault
    '----------------------------------------------------------------
    WindowsXPC1.ColorScheme = XP_OliveGreen '= System ' = XP_Blue
    WindowsXPC1.InitSubClassing
    '----------------------------------------------------------------
    
End Sub

Private Sub cmdCancel_Click()
    NewSelectedImage = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    NewSelectedImage = lvwImages.SelectedItem.Key
    NewSelectedImage = Replace(NewSelectedImage, "K", "", 1, 1)
    Unload Me
    
End Sub

Private Sub tmrLoad_Timer()
    tmrLoad.Enabled = False
    LoadImages
End Sub

Sub LoadImages()

    Dim Files() As String
    Dim imgX As ListImage
    Dim itmX ' As ListItem
    Dim x As Long, Counter As Long
    
    Me.MousePointer = vbHourglass
    
    Me.Caption = fMainForm.Caption
    Me.RightToLeft = fMainForm.RightToLeft
    
    NewSelectedImage = ""
    '------------------------------------------------------
    Files = DirectoryFiles(AppPath + "themes\")
    
    Counter = 0
    For x = LBound(Files) To UBound(Files)
        If InStrRev(LCase(Files(x)), ".jpg") Or InStrRev(LCase(Files(x)), ".gif") Or InStrRev(LCase(Files(x)), ".bmp") Then
            ReDim Preserve ImageFiles(Counter)
            ImageFiles(Counter) = Files(x)
            Counter = Counter + 1
        End If
    Next
    
    imglImages.ImageHeight = 100
    imglImages.ImageWidth = 100
    '------------------------------------------------------
    imglImages.ListImages.Clear
    For x = LBound(ImageFiles) To UBound(ImageFiles)
        imglImages.ListImages.Add x + 1, "K" + ImageFiles(x), LoadPicture(AppPath + "themes\" + ImageFiles(x))
    Next
    '------------------------------------------------------
    Set lvwImages.Icons = imglImages
    
    lvwImages.ListItems.Clear
    For x = LBound(ImageFiles) To UBound(ImageFiles)
        Set itmX = lvwImages.ListItems.Add(, "K" + ImageFiles(x), , x + 1)
        If SelectedImage = ImageFiles(x) Then
            itmX.Selected = True
            Set picView.Picture = LoadPicture(AppPath + "themes\" + ImageFiles(x))
        End If
    Next
    '------------------------------------------------------
    'Set Me.Picture = imglImages.ListImages(3).Picture
    '------------------------------------------------------
    Me.MousePointer = vbDefault
    'lvwImages.ListItems.Item(SelectedImage).Selected = True
    'Set picView.Picture = imglImages.ListImages("K" + CStr(SelectedImage)).Picture

End Sub

Private Sub lvwImages_Click()

    Set picView.Picture = imglImages.ListImages.Item(lvwImages.SelectedItem.index).Picture
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub


