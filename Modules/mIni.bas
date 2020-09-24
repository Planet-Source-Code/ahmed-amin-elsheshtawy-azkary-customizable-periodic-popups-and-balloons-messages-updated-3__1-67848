Attribute VB_Name = "mIni"
'==========================================================
'           Copyright Information
'==========================================================
'Program Name: Azkary
'Program Author   : Elsheshtawy, Ahmed Amin
'Home Page        : http://www.islamware.com
'Copyrights © 2007 Islamware. All rights reserved.
'==========================================================
'Permission to use, copy, modify, and distribute this software and its
'documentation for any purpose and without fee is hereby granted.
'==========================================================
Option Explicit

' See full tutorial at http://www.vbexplorer.com/VBExplorer/focus/ini_tutorial.asp

#If Win16 Then
    Public Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal FileName As String) As Integer
    Public Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal Default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal FileName As String) As Integer
#Else
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Public Function ReadINI(sSection As String, sKeyName As String, sFileName As String) As String
    'Tutorial: http://www.vbexplorer.com/VBExplorer/focus/ini_tutorial.asp
    Dim sRet As String
    sRet = String(5000, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), sFileName))
End Function

Public Function WriteINI(sSection As String, sKeyName As String, ByVal sNewString As String, sFileName) As Long
    'Returns non-zero on success; zero on failure.
    WriteINI = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Public Function GetSectionsINI(sFileName As String) As String()
    'Notice that lReturn contains the value 19 instead of the expected 18.
    'Normally, lReturn indexes the next-to-last character.  However,
    'when retrieving the section names using GetPrivateProfileString or
    'GetProfileString, lReturn indexes the last character.
    'http://www.vbexplorer.com/VBExplorer/focus/ini_tutorial_2.asp
    Dim sRet As String, x As Long
    sRet = String(5000, Chr(0))
    x = GetPrivateProfileString(vbNullString, "", "", sRet, Len(sRet), sFileName)
    If x > 0 Then
        sRet = Left(sRet, x - 1)
    Else
        sRet = ""
    End If
    
    GetSectionsINI = Split(sRet, Chr$(0))
    
End Function

Public Function GetKeysINI(sSection As String, sFileName As String) As String()
    Dim sRet As String
    sRet = ReadINI(sSection, vbNullString, sFileName)
    GetKeysINI = Split(sRet, Chr$(0))
End Function

Public Function DeleteKeyINI(sSection As String, sKeyName As String, sFileName) As Long
    'Returns non-zero on success; zero on failure.
    DeleteKeyINI = WritePrivateProfileString(sSection, sKeyName, vbNullString, sFileName)
End Function

Public Function DeleteSectionINI(sSection As String, sFileName) As Long
    'Returns non-zero on success; zero on failure.
    DeleteSectionINI = WritePrivateProfileString(sSection, vbNullString, "", sFileName)
End Function

