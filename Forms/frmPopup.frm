VERSION 5.00
Begin VB.Form frmPopup 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   -180
      Top             =   90
   End
   Begin VB.Shape PopupShape 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   285
      Left            =   135
      Top             =   45
      Width           =   375
   End
   Begin VB.Label lblMessage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   465
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Effect As AnimeEffectEnum
Public FrameTime As Long
Public FrameCount As Long
Public AnimeEvent As AnimeEventEnum

Private Sub Form_Load()
    
    If Not FrameTime Then FrameTime = 11
    If Not FrameCount Then FrameCount = 33
               
    'MakeTransparent Me
     
    'Me.Move -8000, -8000, 1, 1
    
    'Me.Move 4000, 2000
    'MakeTopMostNoFocus Me.hwnd
    
    'ShowWindow Me.hwnd, SW_SHOWNOACTIVATE
 
    'FadeForm Me.hwnd, 120
    
    'AnimateForm Me, aload, Effect, FrameTime, FrameCount
    
    'FadeForm Me.hWnd, 60
        
    'tmrClose.Enabled = True
    
End Sub

Private Sub Form_Resize()
    
    'MakeTransparent Me
    
    'FadeForm Me.hWnd, 160
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    tmrClose.Enabled = False
    
    If FrameTime < 0 Then FrameTime = 1
    If FrameCount < 1 Then FrameCount = 30
    
    'PreventFocus Me.hwnd
    If Me.WindowState <> vbMinimized Then
        AnimateForm Me, aUnload, Effect, FrameTime, FrameCount
    End If
    
End Sub

Private Sub tmrClose_Timer()
    Unload Me
End Sub
