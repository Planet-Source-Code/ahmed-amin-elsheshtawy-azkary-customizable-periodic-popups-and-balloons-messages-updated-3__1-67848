Attribute VB_Name = "mAutoRun"
Option Explicit

Public Enum eAutoRunTypes
    eNever
    eOnce
    eAlways
End Enum

Public Property Let AutoRun(ByVal eType As eAutoRunTypes)
Dim sExe As String

    sExe = App.Path
    If (Right$(sExe, 1) <> "\") Then sExe = sExe & "\"
    sExe = sExe & App.EXEName
    
    Dim cR As New cRegistry
    cR.ClassKey = HKEY_LOCAL_MACHINE
    If (eType = eNever) Then
        ' Remove entry from always Run if it is there:
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
        cR.ValueKey = App.EXEName
        On Error Resume Next
        cR.DeleteValue
        Err.Clear
        ' Remove entry from RunOnce if it is there:
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
        On Error Resume Next
        cR.DeleteValue
        Err.Clear
    ElseIf eType = eOnce Then
        ' Remove entry from always Run if it is there:
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
        cR.ValueKey = App.EXEName
        On Error Resume Next
        cR.DeleteValue
        Err.Clear
        ' Add an entry to RunOnce (or just ensure the exe name and path
        ' is correct if it is already there):
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
        cR.ValueKey = App.EXEName
        cR.ValueType = REG_SZ
        cR.Value = sExe
    Else
        ' Remove entry from RunOnce if it is there:
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
        cR.ValueKey = App.EXEName
        On Error Resume Next
        cR.DeleteValue
        Err.Clear
        ' Add an entry to RunOnce (or just ensure the exe name and path
        ' is correct if it is already there):
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
        cR.ValueKey = App.EXEName
        cR.ValueType = REG_SZ
        cR.Value = sExe
    End If
        
End Property

Public Property Get AutoRun() As eAutoRunTypes
    Dim cR As New cRegistry
    cR.ClassKey = HKEY_LOCAL_MACHINE
    cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
    cR.ValueKey = App.EXEName
    cR.Default = "?"
    cR.ValueType = REG_SZ
    If (cR.Value = "?") Then
        cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
        If (cR.Value = "?") Then
            AutoRun = eNever
        Else
            AutoRun = eOnce
        End If
    Else
        AutoRun = eAlways
    End If
End Property


