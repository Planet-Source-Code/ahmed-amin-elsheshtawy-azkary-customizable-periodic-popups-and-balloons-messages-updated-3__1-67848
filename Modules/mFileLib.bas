Attribute VB_Name = "mFileLib"
Option Explicit

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileCreatedTime Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, lpCreationTime As FILETIME, ByVal NullLastAccessTime As Long, ByVal NullLastWriteTime As Long) As Long
Private Declare Function SetFileAccessTime Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, ByVal NullCreationTime As Long, lpLastAccessTime As FILETIME, ByVal NullWriteTime As Long) As Long
Private Declare Function SetFileModifiedTime Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, ByVal NullCreationTime As Long, ByVal NullLastAccessTime As Long, lpLastWriteTime As FILETIME) As Long
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

' Convert a SYSTEMTIME into a Date.
Public Function SystemTimeToDate(system_time As SYSTEMTIME) As Date
    With system_time
        SystemTimeToDate = CDate( _
            Format$(.wMonth) & "/" & _
            Format$(.wDay) & "/" & _
            Format$(.wYear) & " " & _
            Format$(.wHour) & ":" & _
            Format$(.wMinute, "00") & ":" & _
            Format$(.wSecond, "00"))
    End With
End Function

' Convert a Date into a SYSTEMTIME.
Public Function DateToSystemTime(ByVal the_date As Date) As SYSTEMTIME
    With DateToSystemTime
        .wYear = Year(the_date)
        .wMonth = Month(the_date)
        .wDay = Day(the_date)
        .wHour = Hour(the_date)
        .wMinute = Minute(the_date)
        .wSecond = Second(the_date)
        '.wDayOfWeek = Weekday(the_date) - 1
    End With
End Function

' Convert the FILETIME structure into a Date.
Public Function FileTimeToDate(file_time As FILETIME) As Date
    Dim system_time As SYSTEMTIME

    ' Convert the FILETIME into a SYSTEMTIME.
    FileTimeToSystemTime file_time, system_time

    ' Convert the SYSTEMTIME into a Date.
    FileTimeToDate = SystemTimeToDate(system_time)
End Function

' Convert a Date into a FILETIME structure.
Public Function DateToFileTime(ByVal the_date As Date) As FILETIME
    Dim system_time As SYSTEMTIME
    Dim file_time As FILETIME

    ' Convert the Date into a SYSTEMTIME.
    system_time = DateToSystemTime(the_date)

    ' Convert the SYSTEMTIME into a FILETIME.
    SystemTimeToFileTime system_time, file_time
    DateToFileTime = file_time
End Function

' Return True if there is an error.
Public Function GetFileTimes(ByVal file_name As String, ByRef creation_date As Date, ByRef access_date As Date, ByRef modified_date As Date, ByVal local_time As Boolean) As Boolean
    
    Dim file_handle As Long
    Dim creation_filetime As FILETIME
    Dim access_filetime As FILETIME
    Dim modified_filetime As FILETIME
    Dim file_time As FILETIME

    ' Assume something will fail.
    GetFileTimes = True

    ' Open the file.
    file_handle = CreateFile(file_name, GENERIC_READ, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        0&, OPEN_EXISTING, 0&, 0&)
    If file_handle = 0 Then Exit Function

    ' Get the times.
    If GetFileTime(file_handle, creation_filetime, access_filetime, modified_filetime) = 0 Then
        CloseHandle file_handle
        Exit Function
    End If

    ' Close the file.
    If CloseHandle(file_handle) = 0 Then Exit Function

    ' See if we should convert to the local
    ' file system time.
    If local_time Then
        ' Convert to local file system time.
        FileTimeToLocalFileTime creation_filetime, file_time
        creation_filetime = file_time

        FileTimeToLocalFileTime access_filetime, file_time
        access_filetime = file_time

        FileTimeToLocalFileTime modified_filetime, file_time
        modified_filetime = file_time
    End If

    ' Convert into dates.
    creation_date = FileTimeToDate(creation_filetime)
    access_date = FileTimeToDate(access_filetime)
    modified_date = FileTimeToDate(modified_filetime)

    GetFileTimes = False
End Function

' Return True if there is an error.
Public Function SetFileTimes(ByVal file_name As String, ByVal creation_date As Date, ByVal access_date As Date, ByVal modified_date As Date, ByVal local_times As Boolean) As Boolean
    
    Dim file_handle As Long
    Dim creation_filetime As FILETIME
    Dim access_filetime As FILETIME
    Dim modified_filetime As FILETIME
    Dim file_time As FILETIME

    ' Assume something will fail.
    SetFileTimes = True

    ' Convert the dates into FILETIMEs.
    creation_filetime = DateToFileTime(creation_date)
    access_filetime = DateToFileTime(access_date)
    modified_filetime = DateToFileTime(modified_date)

    ' Convert the file times into system file times.
    If local_times Then
        LocalFileTimeToFileTime creation_filetime, file_time
        creation_filetime = file_time

        LocalFileTimeToFileTime access_filetime, file_time
        access_filetime = file_time

        LocalFileTimeToFileTime modified_filetime, file_time
        modified_filetime = file_time
    End If

    ' Open the file.
    file_handle = CreateFile(file_name, GENERIC_READ, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        0&, OPEN_EXISTING, 0&, 0&)
    If file_handle = 0 Then Exit Function

    ' Set the times.
    If SetFileTime(file_handle, creation_filetime, access_filetime, modified_filetime) = 0 Then
        CloseHandle file_handle
        Exit Function
    End If

    ' Close the file.
    If CloseHandle(file_handle) = 0 Then Exit Function

    SetFileTimes = False
End Function

' Return True if there is an error.
Public Function SetFileModifiedDate(ByVal file_name As String, ByVal modified_date As Date, ByVal local_times As Boolean) As Boolean
    
    Dim file_handle As Long
    Dim modified_filetime As FILETIME
    Dim file_time As FILETIME

    ' Assume something will fail.
    SetFileModifiedDate = True

    ' Convert the date into a FILETIME.
    modified_filetime = DateToFileTime(modified_date)

    ' Convert the file time into a system file time.
    If local_times Then
        LocalFileTimeToFileTime modified_filetime, file_time
        modified_filetime = file_time
    End If

    ' Open the file.
    file_handle = CreateFile(file_name, GENERIC_WRITE, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        0&, OPEN_EXISTING, 0&, 0&)
    If file_handle = 0 Then Exit Function

    ' Set the time.
    If SetFileModifiedTime(file_handle, ByVal 0&, ByVal 0&, modified_filetime) = 0 Then
        CloseHandle file_handle
        Exit Function
    End If

    ' Close the file.
    If CloseHandle(file_handle) = 0 Then Exit Function

    SetFileModifiedDate = False
End Function

' Return True if there is an error.
Public Function SetFileAccessedDate(ByVal file_name As String, ByVal accessed_date As Date, ByVal local_times As Boolean) As Boolean
    Dim file_handle As Long
    Dim accessed_filetime As FILETIME
    Dim file_time As FILETIME

    ' Assume something will fail.
    SetFileAccessedDate = True

    ' Convert the date into a FILETIME.
    accessed_filetime = DateToFileTime(accessed_date)

    ' Convert the file time into a system file time.
    If local_times Then
        LocalFileTimeToFileTime accessed_filetime, file_time
        accessed_filetime = file_time
    End If

    ' Open the file.
    file_handle = CreateFile(file_name, GENERIC_WRITE, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        0&, OPEN_EXISTING, 0&, 0&)
    If file_handle = 0 Then Exit Function

    ' Set the time.
    If SetFileAccessTime(file_handle, ByVal 0&, accessed_filetime, ByVal 0&) = 0 Then
        CloseHandle file_handle
        Exit Function
    End If

    ' Close the file.
    If CloseHandle(file_handle) = 0 Then Exit Function

    SetFileAccessedDate = False
End Function

' Return True if there is an error.
Public Function SetFileCreatedDate(ByVal file_name As String, ByVal created_date As Date, ByVal local_times As Boolean) As Boolean
    
    Dim file_handle As Long
    Dim created_filetime As FILETIME
    Dim file_time As FILETIME

    ' Assume something will fail.
    SetFileCreatedDate = True

    ' Convert the date into a FILETIME.
    created_filetime = DateToFileTime(created_date)

    ' Convert the file time into a system file time.
    If local_times Then
        LocalFileTimeToFileTime created_filetime, file_time
        created_filetime = file_time
    End If

    ' Open the file.
    file_handle = CreateFile(file_name, GENERIC_WRITE, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        0&, OPEN_EXISTING, 0&, 0&)
    If file_handle = 0 Then Exit Function

    ' Set the time.
    If SetFileCreatedTime(file_handle, created_filetime, ByVal 0&, ByVal 0&) = 0 Then
        CloseHandle file_handle
        Exit Function
    End If

    ' Close the file.
    If CloseHandle(file_handle) = 0 Then Exit Function

    SetFileCreatedDate = False
End Function
