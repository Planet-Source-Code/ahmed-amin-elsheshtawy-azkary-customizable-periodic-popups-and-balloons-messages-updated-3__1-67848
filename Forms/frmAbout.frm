VERSION 5.00
Object = "{C151518A-D64D-4C66-96F3-DB69BF286B30}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":058A
   ScaleHeight     =   3930
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "About prjSnapShotMgr"
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -1665
      Top             =   3420
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   240
      Left            =   240
      Picture         =   "frmAbout.frx":7EE1
      ScaleHeight     =   240
      ScaleMode       =   0  'User
      ScaleWidth      =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   240
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2160
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   3465
      Width           =   1467
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   4185
      Picture         =   "frmAbout.frx":846B
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyrights (c) Islamware Corporation. All rights reserved."
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Website address:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1380
      Width           =   1275
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Support Email:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Emails:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblMewsoft 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.islamware.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1560
      MouseIcon       =   "frmAbout.frx":8BD5
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1380
      Width           =   2550
   End
   Begin VB.Label lblSupportEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "support@islamware.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1575
      MouseIcon       =   "frmAbout.frx":8EDF
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1680
      Width           =   2580
   End
   Begin VB.Label lblSalesEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "sales@islamware.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1560
      MouseIcon       =   "frmAbout.frx":91E9
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2040
      Width           =   2580
   End
   Begin VB.Image Image4 
      Height          =   360
      Left            =   7620
      Picture         =   "frmAbout.frx":94F3
      Top             =   600
      Width           =   360
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   4230
      Picture         =   "frmAbout.frx":9C5D
      Top             =   1260
      Width           =   480
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Azkary"
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1080
      TabIndex        =   4
      Tag             =   "Application Title"
      Top             =   120
      Width           =   4725
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   270
      X2              =   5805
      Y1              =   2385
      Y2              =   2385
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   270
      X2              =   5805
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   225
      Left            =   1080
      TabIndex        =   3
      Tag             =   "Version"
      Top             =   600
      Width           =   4725
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":AA9F
      ForeColor       =   &H00000000&
      Height          =   810
      Left            =   255
      TabIndex        =   2
      Tag             =   "Warning: ..."
      Top             =   2565
      Width           =   5625
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = gAppName
    Me.Caption = "About " & gAppName
    '----------------------------------------------------------------
    WindowsXPC1.ColorScheme = XP_Blue
    WindowsXPC1.InitSubClassing
    '----------------------------------------------------------------
    
End Sub

Private Sub lblMewsoft_Click()
    Dim URL As String
    URL = Join(Array("h", "t", "t", "p", ":", "/", "/", "w", "w", "w", ".", "i", "s", "l", "a", "m", "w", "a", "r", "e", ".", "c", "o", "m"), "~")
    URL = Replace(URL, "~", "")
    ShellDocument URL
End Sub

Private Sub lblSalesEmail_Click()
    Dim URL As String
    URL = Join(Array("m", "a", "i", "l", "t", "o", ":", "s", "a", "l", "e", "s", "@", "i", "s", "l", "a", "m", "w", "a", "r", "e", ".", "c", "o", "m"), "")
    URL = Replace(URL, "~", "")
    ShellDocument URL
End Sub

Private Sub lblSupportEmail_Click()
    Dim URL As String
    URL = Join(Array("m", "a", "i", "l", "t", "o", ":", "s", "u", "p", "p", "o", "r", "t", "@", "i", "s", "l", "a", "m", "w", "a", "r", "e", ".", "c", "o", "m"), "")
    URL = Replace(URL, "~", "")
    ShellDocument URL
End Sub

Private Sub cmdOK_Click()
        Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

