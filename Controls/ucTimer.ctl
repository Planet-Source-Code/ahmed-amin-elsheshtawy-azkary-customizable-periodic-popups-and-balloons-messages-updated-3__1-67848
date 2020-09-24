VERSION 5.00
Begin VB.UserControl ucTimer 
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1155
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ucTimer.ctx":0000
   ScaleHeight     =   930
   ScaleWidth      =   1155
   ToolboxBitmap   =   "ucTimer.ctx":0382
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   585
      Top             =   135
   End
End
Attribute VB_Name = "ucTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim TotalElapsed  As Double
Dim m_lInterval As Long
Dim m_vTag As Variant

Public Event Timer()

Private Sub Timer1_Timer()
    
    TotalElapsed = TotalElapsed + 1
    If TotalElapsed >= m_lInterval Then
        TotalElapsed = 0
        RaiseEvent Timer
    End If
    
End Sub

Private Sub UserControl_Initialize()
    TotalElapsed = 0
    m_lInterval = 1
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 17 * 15 '420
    UserControl.Height = 16 * 15 '420
End Sub

Public Property Get Enabled() As Boolean
    Enabled = Timer1.Enabled
End Property

Public Property Let Enabled(bEnabled As Boolean)
    TotalElapsed = 0
    Timer1.Enabled = bEnabled
End Property

Public Property Get Interval() As Double
    Interval = m_lInterval
End Property

Public Property Let Interval(lInterval As Double)
    m_lInterval = lInterval
End Property

Property Get Tag() As Variant
    Tag = m_vTag
End Property

Property Let Tag(varTag As Variant)
    m_vTag = varTag
End Property


Public Property Get Elapsed() As Long
    Elapsed = TotalElapsed
End Property

Private Sub UserControl_Terminate()
    'Timer1.Enabled = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", Me.Enabled, 0
    PropBag.WriteProperty "Interval", Me.Interval, 0
    PropBag.WriteProperty "Tag", Me.Tag, ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.Enabled = PropBag.ReadProperty("Enabled", 0)
    Me.Interval = PropBag.ReadProperty("Interval", 0)
    Me.Tag = PropBag.ReadProperty("Tag", "")
End Sub


