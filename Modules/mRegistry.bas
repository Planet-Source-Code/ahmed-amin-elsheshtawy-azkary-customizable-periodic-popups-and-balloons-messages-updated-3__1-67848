Attribute VB_Name = "mRegistry"
            '--------------------------------------------[-][x]--'
            '----------------[ Registry Access ]-----------------'
            '----------------------------------------------------'
            '     ______________________                         '
            '     Name: clsRegistryAccess                        '
            '     Source type: Class Module                      '
            '     Version: 2.05                                  '
            '     Author: Aleksandar Ruzicic a.k.a. krckoorascic '
            '     Modified by: <nobady so far>                   '
            '     Contact: krckoorascic@gmail.com                '
            '     Last update: Friday, April 08, 2005 20:51      '
            '     ______________________                         '
            '     BIG Thanx goes to:                             '
            '     -The KPD-Team for their API-Guide!             '
            '     -Chirstoph von Wittich for his ApiViewer 2004  '
            '     -mladenovicz & Shadowed for all their help     '
            '     -and all folks from EliteSecurity.org          '
'-------------------------------------------------------------------------------'
'[-]Licence:                                                                    '
'You are free to use this class in the way you like it, but it will be nice to  '
'to put me in credits ;o). If you modify something (fix bug or add new features)'
'please put your name above (in 'Modified by') and mail me.                     '
'                                                                               '
'[-]Contents:                                                                   '
' [1] - rcMainKey (Enum)                                                        '
' [2] - rcRegType (Enum)                                                        '
' [3] - CreateKeyIfDoesntExists (Property)                                      '
' [4] - GetKeys (Private Function)                                              '
' [5] - CreateKey (Function)                                                    '
' [6] - KillKey (Function)                                                      '
' [7] - KeyExists (Function)                                                    '
' [8] - EnumKeys (Function)                                                     '
' [9] - HaveSubKey (Function)                                                   '
' [10] - WriteString (Function)                                                 '
' [11] - ReadString (Function)                                                  '
' [12] - WriteDWORD (Function)                                                  '
' [13] - ReadDWORD (Function)                                                   '
' [14] - WriteBinary (Function)                                                 '
' [15] - ReadBinary (Function)                                                  '
' [16] - KillValue (Function)                                                   '
' [17] - ValueExists (Function)                                                 '
' [18] - EnumValues (Function)                                                  '
' [19] - ExportToReg (Function)                                                 '
' [20] - generateReg (Private Function)                                         '
' [21] - ImportFromReg Function)                                                '
' [22] - StrToBin (Private Function)                                            '
' [23] - BinToStr (Private Function)                                            '
' [24] - isBinValid (Function)                                                  '
' _____________________________________                                         '
' [1] rcMainKey (Enum)                                                          '
' Enum (public) which holds all root keys (hkeys). Its only possible to write to'
' first five: HKCR, HKCU, HKLM, HKUS, HKPD.                                     '
'                                                                               '
' [2] rcRegType (Enum)                                                          '
' Enum (Public) holds allregistry data types. In this class only tree main are  '
' covered, that are: REG_SZ, REG_BINARY and REG_DWORD.                          '
'                                                                               '
' [3] CreateKeyIfDoesntExists (Property)                                        '
' Property, Boolean, Let/Get; when writting some value to registry, in some key,'
' if that key doesnt exist then it will be created if this property is true, but'
' if is set to false then writting function will return error (0).              '
'                                                                               '
' [4] GetKeys (Private Function)                                                '
' Used to separate given path (to some key) in two values, the hkey value and to'
' subkey of that hkey. This function also alows you to use short constants, e.g.'
' for:                                                                          '
'       HKEY_LOCAL_MACHINE                                                      '
' you may use:                                                                  '
'       HKLM                                                                    '
' which is very simpler.                                                        '
' This is the list of short consts:                                             '
'       HKEY_CLASSES_ROOT ........ HKCR                                         '
'       HKEY_CURRENT_USER ........ HKCU                                         '
'       HKEY_LOCAL_MACHINE ....... HKLM                                         '
'       HKEY_USERS ............... HKUS                                         '
'       HKEY_PERFORMANCE_DATA .... HKPD                                         '
'       HKEY_CURRENT_CONFIG ...... HKCC                                         '
'       HKEY_DYN_DATA ............ HKDD                                         '
'                                                                               '
' [5] CreateKey (Function)                                                      '
' Creates new key in Registry                                                   '
' CreateKey(sPath) As Long                                                      '
' sPath - string; path to the key to create                                     '
'       CreateKey("HKCU\Software\ES")                                           '
' Function returns:                                                             '
' 0 - if there is an error                                                      '
' handle of created key, non-zero, if success                                   '
'                                                                               '
' [6] KillKey (Function)                                                        '
' Deletes existing key (and all of his subkeys) from Registry                   '
' KillKey(sPath) As Long                                                        '
' sPath - string; path to the key to delete                                     '
'       KillKey("HKCU\Software\ES")                                             '
' Function returns::                                                            '
' 0 - if there is an error                                                      '
' handle deleted key, non-zero, if success                                      '
'                                                                               '
' [7] KeyExists (Function)                                                      '
' Cheks if key exists.                                                          '
' KeyExists(sPath) As Boolean                                                   '
' sPath - string; path to the key to check                                      '
'        KeyExists("HKCU\Software\ES\Login")                                    '
' Function returns:                                                             '
' True - key exists                                                             '
' False - key doesn't exists                                                    '
'                                                                               '
' [8] EnumKeys (Function)                                                       '
' Returns an array of all subkeys of given key.                                 '
' EnumKeys(sPath As String, Key() As String) As Long                            '
' sPath - string; path to the key which subkeys will be returned                '
' Key() - string array; empty array which will hold names of subkeys            '
'       EnumKeys("HKCU\Software",Ime)                                           '
' Function returns:                                                             '
' -1 - if there is an error                                                     '
' number of subkeys if success                                                  '
' filled zero-based string array (filled with names)                            '
'                                                                               '
' [9] HaveSubKey (Function)                                                     '
' Returns true if given key have atleast one key                                '
' HaveSubKey(sPathAs String) As Boolean                                         '
' sPath - string; path to the key which will be checked for subkeys             '
'       HaveSubKeys("HKCU\Software\ES")                                         '
' Function returns:                                                             '
' true - in given key exists atleast one subkey                                 '
' false - no subkeys in this key                                                '
'                                                                               '
' [10] WriteString (Function)                                                   '
' Writes/edits string data type into registry                                   '
' WriteString(sPath, sName, sValue) As Long                                     '
' sPath - string; path to key                                                   '
' sName - string; name of data                                                  '
' sValue - string; data value                                                   '
'       WriteString("HKCU\Software\ES", "@", "http://www.elitesecurity.org")    '
' NOTE: for editing '(Default)' data, as a name you may use '@' or simply left  '
' the name parametar empty - "" (vbNullString)                                  '
' Function returns:                                                             '
' 0 - if fails                                                                  '
' handle of key in which is writed string data, if success                      '
'                                                                               '
' [11] ReadString (Function)                                                    '
' Reads string data from registry                                               '
' ReadString(sPath, sName, [Default]) As String                                 '
' sPath - string; path to the key                                               '
' sName - string; name of data to read                                          '
' sDefault - string; optional parametar which will be returned if function fails'
'       ReadString("HKCU\Software\ES", "Username", "krckoorascic")              '
' Function returns:                                                             '
' sDefault parametar - if fails (maybe data or key doesn't exists)              '
' string value if success                                                       '
'                                                                               '
' [12] WriteDWORD (Function)                                                    '
' Writes/edits DWORD data type into registry                                    '
' WriteDWORD(sPath, sName, lValue) As Long                                      '
' sPath - string; path to the key                                               '
' sName - string; name of data                                                  '
' lValue - long; value to be stored                                             '
'       WriteDWORD("HKCU\Software\ES", "AutoLogin", 1)                          '
' Function returns:                                                             '
' 0 - if there is an error                                                      '
' handle of key in which is writted dword data, if success                      '
'                                                                               '
' [13] ReadDWORD (Function)                                                     '
' Reads dword data type from registry                                           '
' ReadDWORD(sPath, sName, [lDefault]) As Long                                   '
' sPath - string; path to the key                                               '
' sName - string; name of data                                                  '
' lDefault - long; optional parametar (default -1) which will be returned if    '
'   function fails.                                                             '
'       ReadDWORD("HKCU\Software\ES", "AutoLogin", 0)                           '
' Function returns:                                                             '
' lDefault parametar - on error                                                 '
' long value - if success                                                       '
'                                                                               '
' [14] WriteBinary (Function)                                                   '
' Writes/edits Binary data type into registry                                   '
' WriteBinary(sPath, sName, sValue) As Long                                     '
' sPath - string; path to the key                                               '
' sName - string; name of data                                                  '
' sValue - string; value that will be stored into registry. It MUST be in HEX   '
' format, it's not needed to be uppercase and space after each two chars is not '
' needed (anything except "A-F", "0-9" i " "[space] is NOT valid!!)             '
'       WriteBinary("HKCU\Software\ES", "Password", "FF 20 3E 0B AF 00 00")     '
' Function returns:                                                             '
' 0 - if there is an error                                                      '
' handle of key in which is writted dword data, if success                      '
'                                                                               '
' [15] ReadBinary (Function)                                                    '
' Reads binary data type from registry                                          '
' ReadBinary(sPath, sName, [sDefault]) As String                                '
' sPath - string; path to the key                                               '
' sName - string; name of data                                                  '
' sDefault - string; optional (vbNullChar - Chr$(0)) which is returned if error '
' occurs...                                                                     '
'        ReadBinary("HKCU\Software\ES", "Password", "FF 20 3E 0B AF 00 00")     '
' Function returns:                                                             '
' sDefault - if there is an error                                               '
' string value if success                                                       '
'                                                                               '
' [16] KillValue (Function)                                                     '
' Deletes any type data from registry                                           '
' KillValue(sPath, sName) As Long                                               '
' sPath - string; path to the key                                               '
' sName - string; name of data                                                  '
'        KillValue("HKCU\Software\ES", "Password")                              '
' Function returns:                                                             '
' 0 - if there is an error (value is NOT deleted)                               '
' handle of key where we killed the value...                                    '
'                                                                               '
' [17] ValueExists (Function)                                                   '
' Checks if some value (of any type) exists in registry database                '
' ValueExists(sPath, sName) As Boolean                                          '
' sPath - string; path to the key                                               '
' sName - string; name of value that should be checked                          '
'       ValueExists("HKCU\Software\ES", "Username")                             '
' Function returns:                                                             '
' True - if value exists in given key                                           '
' False - if not exists                                                         '
'                                                                               '
' [18] EnumValues (Function)                                                    '
' Returns arrays of names and values for all data in specified key              '
' EnumValues(sPath, sName(), sValue(), [OnlyType]) As Long                      '
' sPath - string; path to the key                                               '
' sName() - array (string) that will be populated with names of data            '
' sValue() - array (variant) that will be populated with data values            '
' OnlyType - rcRegType, optional parametar (REG_NONE - 0) which is filter for   '
' reading values (if OnlyType = REG_SZ, only data of string type will be readed)'
' if this param is not given (REG_NONE) then all three data types are returned  '
' NOTE: none of the arrays is NOT sorted.                                       '
'        EnumValues("HKCU\Software", Ime, Vrednost, REG_BINARY)                 '
' Function returns:                                                             '
' -1 - if there is an error                                                     '
' number of readed values (if success)                                          '
' filled arrays (of names and values) that are 0-based...                       '
'                                                                               '
' [19] ExportToReg (Function)                                                   '
' Generate .reg fole (same as Windows Registry Editor - Regedit)                '
' ExportToReg(sPath, sRegFile [IncludeSubkeys], [Output]) As Long               '
' sPath - string; path to the key where exporting starts                        '
' sRegFile - string; path to the file that will be generated (if file allready  '
' exists then 0 is returned - error)                                            '
' Output - TextBox object, optional. If TextBox is given then the name of key   '
' thats currently reading will be displayed in this TextBox object (i added this'
' feature cuz exporting of realy large keys (like root keys) on slow machines   '
' may take very long time to finish, you may use (use Change() event of textbox)'
' this to show user some progress...                                            '
' IncludeSubkeys - boolean, optional (True). If set to false only contents of   '
' given key will be exported, otherwise it will also return all values from all '
' subkeys of given key                                                          '
'        ExportToReg("HKCU\Software\ES", "C:\ES.reg")                           '
' Function returns:                                                             '
' 0 - if there is an error                                                      '
' 1 - if .reg file is successufuly generated                                    '
'                                                                               '
' [20] generateReg (Private Function)                                           '
' Private (recurzive) function that does real exporting(called from ExportToReg)'
' Function returns:                                                             '
' False - if there is an error                                                  '
' True - if key is readed successufuly (and written to file)                    '
'                                                                               '
' [21] ImportFromReg (Function)                                                 '
' Imports .reg file in registry database (same as Regedit)                      '
' ImportFromReg(sRegFile) As Long                                               '
' sRegFile - string; path to the .reg file                                      '
'       ImportFromReg("C:\ES.reg")                                              '
' Function returns:                                                             '
' 0 - if there is an error (or file not exists)                                 '
' 1 - if success                                                                '
'                                                                               '
' [22] StrToBin (Private Function)                                              '
' Used for WriteBinary function. Converts i.e "BE 3E FF AB" into "¾>ÿ«" (value  '
' that will be converted into byte array an recorded to registry)               '
'                                                                               '
' [23] BinToStr (Private Function)                                              '
' Used for ReadBinary function. Converts i.e "¾>ÿ«" into "BE 3E FF AB" (value in'
' human-readable format)                                                        '
'                                                                               '
' [24] isBinValid (Function)                                                    '
' Checks if given value is in valid hex format (used for WriteBinary function). '
' isBinValid(sBin) As Boolean                                                   '
' sBin - string; that will be checked for validability                          '
'        isBinValid("3E BE 00 AS") - ovde ce vratiti False                      '
' Function returns:                                                             '
' True - if string (sBin) don't contains anything except "A"-"F" 0-9 i " " space'
' False - if string is not in valid hex format...                               '
'                                                                               '
'-------------------------------------------------------------------------------'
'          Copyright © 2004-2005, krckoorascic, krckoorascic@gmail.com          '
'-------------------------------------------------------------------------------'

Option Explicit

'----[ API's ]----'
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'----[ Constants ]----'
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1009&
Private Const ERROR_BADKEY = 1010&
Private Const ERROR_CANTOPEN = 1011&
Private Const ERROR_CANTREAD = 1012&
Private Const ERROR_CANTWRITE = 1013&
Private Const ERROR_OUTOFMEMORY = 14&
Private Const ERROR_INVALID_PARAMETER = 87&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_MORE_DATA = 234&
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const KEY_CREATE_LINK = &H20&
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Private Const KEY_EXECUTE = KEY_READ
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

'----[ Enums ]----'
Public Enum rcMainKey       'root keys constants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Public Enum rcRegType       'data types constants
    REG_NONE = 0
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7
    REG_RESOURCE_LIST = 8
    REG_FULL_RESOURCE_DESCRIPTOR = 9
    REG_RESOURCE_REQUIREMENTS_LIST = 10
End Enum

'----[ Dim's ]----'
Private hKey             As Long
Private mainKey          As Long
Private sKey             As String
Private lBufferSize      As Long
Private lDataSize        As Long
Private ByteArray()      As Byte
Private createNoExists   As Boolean

'----[ CreateKeyIfDoesntExists ]----'
Public Property Let CreateKeyIfDoesntExists(ByVal offon As Boolean)
    createNoExists = offon
End Property
Public Property Get CreateKeyIfDoesntExists() As Boolean
    CreateKeyIfDoesntExists = createNoExists
End Property
'----[ GetKeys ]----'
Private Function GetKeys(sPath As String, sKey As String) As rcMainKey
Dim pos As Long, mk As String
    
    'replace long with short root constants
    sPath = Replace$(sPath, "HKEY_CURRENT_USER", "HKCU", , , 1)
    sPath = Replace$(sPath, "HKEY_LOCAL_MACHINE", "HKLM", , , 1)
    sPath = Replace$(sPath, "HKEY_CLASSES_ROOT", "HKCR", , , 1)
    sPath = Replace$(sPath, "HKEY_USERS", "HKUS", , , 1)
    sPath = Replace$(sPath, "HKEY_PERFORMANCE_DATA", "HKPD", , , 1)
    sPath = Replace$(sPath, "HKEY_DYN_DATA", "HKDD", , , 1)
    sPath = Replace$(sPath, "HKEY_CURRENT_CONFIG", "HKCC", , , 1)
    
    pos = InStr(1, sPath, "\") 'get pos of first slash

    If (pos = 0) Then 'writting to root
        mk = UCase$(sPath)
        sKey = ""
    Else
        mk = UCase$(Left$(sPath, 4)) 'get hkey
        sKey = Right$(sPath, Len(sPath) - pos) 'get path
    End If
    
    Select Case mk 'return main key handle
        Case "HKCU": GetKeys = HKEY_CURRENT_USER
        Case "HKLM": GetKeys = HKEY_LOCAL_MACHINE
        Case "HKCR": GetKeys = HKEY_CLASSES_ROOT
        Case "HKUS": GetKeys = HKEY_USERS
        Case "HKPD": GetKeys = HKEY_PERFORMANCE_DATA
        Case "HKDD": GetKeys = HKEY_DYN_DATA
        Case "HKCC": GetKeys = HKEY_CURRENT_CONFIG
    End Select
    
End Function
'----[ CreateKey ]----'
Public Function CreateKey(ByVal sPath As String) As Long
    
    hKey = GetKeys(sPath, sKey) 'get keys
    
    'try to create key
    If (RegCreateKey(hKey, sKey, mainKey) = ERROR_SUCCESS) Then
        RegCloseKey mainKey
        CreateKey = mainKey 'success
    Else
        CreateKey = 0 'error
    End If

End Function
'----[ KillKey ]----'
Public Function KillKey(ByVal sPath As String) As Long

    hKey = GetKeys(sPath, sKey)
    
    'try to delete key
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_ALL_ACCESS, mainKey) = ERROR_SUCCESS) Then
        RegDeleteKey mainKey, ""   'delete key
        RegCloseKey mainKey
        KillKey = mainKey 'success
    Else
        KillKey = 0 'error
    End If

End Function
'----[ KeyExists ]----'
Public Function KeyExists(ByVal sPath As String) As Boolean

    hKey = GetKeys(sPath, sKey)
    
    'try to open key
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_ALL_ACCESS, mainKey) = ERROR_SUCCESS) Then
        KeyExists = True 'if we open it than it exists ;o)
        RegCloseKey mainKey 'close key
    Else
        KeyExists = False ' noup, the key don't exists
    End If

End Function
'----[ EnumKeys ]----'
Public Function EnumKeys(ByVal sPath As String, Key() As String) As Long
    Dim sName As String, RetVal As Long
    
    hKey = GetKeys(sPath, sKey)
    
    Erase Key 'clear array
    
    'try to open key
    If (RegOpenKey(hKey, sKey, mainKey) = ERROR_SUCCESS) Then

        EnumKeys = 0 'array is 0-based
        sName = Space(255)
        RetVal = Len(sName)
        
        'do while not we get ERROR_NO_MORE_ITEMS
        While RegEnumKeyEx(mainKey, EnumKeys, sName, RetVal, ByVal 0&, _
                           vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            
            ReDim Preserve Key(EnumKeys) 'incermenting array (+1)
            
            Key(EnumKeys) = Left$(sName, RetVal) 'adding key name to array
                        
            'pripering values for next data
            EnumKeys = EnumKeys + 1 'incerment the counter
            sName = Space(255)
            RetVal = Len(sName)
            
        Wend 'looping ;o)
    
        RegCloseKey mainKey 'close the key
    Else
        EnumKeys = -1 'error (key doesn't exists)
    End If
    
End Function
'----[ HaveSubkey ]----'
Public Function HaveSubkey(ByVal sPath As String) As Boolean
    Dim sName As String, RetVal As Long, SubKeyCount As Long
    
    hKey = GetKeys(sPath, sKey)
    
    If (RegOpenKey(hKey, sKey, mainKey) = ERROR_SUCCESS) Then 'try to open key

        SubKeyCount = 0
        sName = Space(255)
        RetVal = Len(sName)
        HaveSubkey = False
        
        Do While RegEnumKeyEx(mainKey, SubKeyCount, sName, RetVal, ByVal 0&, _
                           vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            'will exit on first loop (we found 1 subkey!)
            HaveSubkey = True
            Exit Do
        Loop
    
        RegCloseKey mainKey 'close the key
    Else
        HaveSubkey = False 'no subkeys in this key
    End If
    
End Function
'----[ WriteString ]----'
Public Function WriteString(ByVal sPath As String, ByVal sName As String, _
                                                   ByVal sValue As String) As Long
                            

    If (KeyExists(sPath) = False) Then 'if key don't exists,
        If (createNoExists = True) Then 'and if CreateKeyIfDoesntExists = True
            CreateKey sPath  ' then create it ;o)
        Else
            WriteString = 0 'error!
            Exit Function
        End If
    End If
    
    hKey = GetKeys(sPath, sKey) 'parse keys
    
    If (sName = "@") Then sName = "" '(Default)
    
    'try to open key
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_WRITE, mainKey) = ERROR_SUCCESS) Then
        'try to write data
        If (RegSetValueEx(mainKey, sName, 0, REG_SZ, ByVal sValue, Len(sValue)) = ERROR_SUCCESS) Then
            RegCloseKey mainKey 'close key
            WriteString = mainKey 'success!
        Else
            WriteString = 0 'error writting data
      End If
    Else
         WriteString = 0 'error opening key
    End If

End Function
'----[ ReadString ]----'
Public Function ReadString(ByVal sPath As String, ByVal sName As String, _
                           Optional sDefault As String = vbNullChar) As String
    
    Dim sData As String, lDuz As Long
    
    hKey = GetKeys(sPath, sKey)
    
    'try to open key
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_READ, mainKey) = ERROR_SUCCESS) Then
        sData = Space(255)     'make buffer
        lDuz = Len(sData)      'get buffer size (255)
        'try to query data
        If (RegQueryValueEx(mainKey, sName, 0, REG_SZ, sData, lDuz) = ERROR_SUCCESS) Then
            RegCloseKey mainKey 'close key
            sData = Trim$(sData) 'trims string
            ReadString = Left$(sData, Len(sData) - 1) 'returning readed value
        Else
            ReadString = sDefault 'return default value (error)
        End If
    Else
        ReadString = sDefault 'return default value (error)
    End If

End Function
'----[ WriteDWORD ]----'
Public Function WriteDWORD(ByVal sPath As String, ByVal sName As String, _
                                                  ByVal lValue As Long) As Long

    If (KeyExists(sPath) = False) Then
        If (createNoExists = True) Then
            CreateKey sPath
        Else
            WriteDWORD = 0
            Exit Function
        End If
    End If

    hKey = GetKeys(sPath, sKey)

    If (RegOpenKeyEx(hKey, sKey, 0, KEY_WRITE, mainKey) = ERROR_SUCCESS) Then
        'try to write data
        If (RegSetValueExA(mainKey, sName, 0, REG_DWORD, lValue, 4) = ERROR_SUCCESS) Then
            RegCloseKey mainKey
            WriteDWORD = mainKey 'yeap, we did it! ;o)
        Else
            WriteDWORD = 0 'error :(
        End If
    Else
        WriteDWORD = 0 'error
    End If

End Function
'----[ ReadDWORD ]----'
Public Function ReadDWORD(ByVal sPath As String, ByVal sName As String, _
                         Optional lDefault As Long = -1) As Long
    Dim lData As Long
    
    hKey = GetKeys(sPath, sKey)
    
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_READ, mainKey) = ERROR_SUCCESS) Then
        'read data
        If (RegQueryValueExA(mainKey, sName, 0, REG_DWORD, lData, 4) = ERROR_SUCCESS) Then
            RegCloseKey mainKey 'close key
            ReadDWORD = lData 'and return readed data
        Else
            ReadDWORD = lDefault 'return default value (data don't exists)
        End If
    Else
        ReadDWORD = lDefault 'return defalt value (key don't exists)
    End If

End Function
'----[ WriteBinary ]----'
Public Function WriteBinary(ByVal sPath As String, ByVal sName As String, _
                                                   ByVal sValue As String) As Long
    Dim l As Long, lDuz As Long, B() As Byte
    
    If (KeyExists(sPath) = False) Then
        If (createNoExists = True) Then
            CreateKey sPath
        Else
            WriteBinary = 0
            Exit Function
        End If
    End If

    hKey = GetKeys(sPath, sKey)
    
    '"translating" value
    sValue = StrToBin(sValue)
   
   'try to open key
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_WRITE, mainKey) = ERROR_SUCCESS) Then
      
        lDuz = Len(sValue) 'get length
        ReDim B(lDuz) As Byte 'redimension array
      
        For l = 1 To lDuz 'making byte array from ascii values for each char
            B(l) = Asc(Mid$(sValue, l, 1))
        Next
        
        If (lDuz = 0) Then ' (zero-length binary value)
            ReDim B(1) As Byte 'we need only array item with index 1
            B(1) = 0
        End If
        
        'try to write data
        If (RegSetValueExB(mainKey, sName, 0, REG_BINARY, B(1), lDuz) = ERROR_SUCCESS) Then
            RegCloseKey mainKey 'close key
            WriteBinary = mainKey 'success
        Else
            WriteBinary = 0 'error
        End If
    Else
         WriteBinary = 0 'key don't exists
    End If

End Function
'----[ ReadBinary ]----'
Public Function ReadBinary(ByVal sPath As String, ByVal sName As String, _
                           Optional sDefault As String = vbNullString) As String
    
    Dim lDuz As Long, sData As String
    
    hKey = GetKeys(sPath, sKey)
    
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_READ, mainKey) = ERROR_SUCCESS) Then
        lDuz = 1 'length
        RegQueryValueEx mainKey, sName, 0, REG_BINARY, 0, lDuz 'get size of data
        sData = Space(lDuz) 'make string buffer
        'now read data
        If (RegQueryValueEx(mainKey, sName, 0, REG_BINARY, sData, lDuz) = ERROR_SUCCESS) Then
            RegCloseKey mainKey 'close key
            ReadBinary = Trim$(BinToStr(sData)) 'trim and convert value
        Else
            ReadBinary = sDefault 'return default value (error)
        End If
    Else
        ReadBinary = sDefault 'return default value (error)
    End If

End Function
'----[ KillValue ]----'
Public Function KillValue(ByVal sPath As String, ByVal sName As String) As Long

    hKey = GetKeys(sPath, sKey) 'parse keys
    
    'open key
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_READ, mainKey) = ERROR_SUCCESS) Then
        RegDeleteValue mainKey, sName 'delete it
        RegCloseKey mainKey 'then close mainkey
        KillValue = mainKey 'and returns it (success)
    Else
        KillValue = 0 'we faild :(
    End If
    
End Function
'----[ ValueExists ]----'
Public Function ValueExists(ByVal sPath As String, ByVal sName As String) As Boolean
    
    hKey = GetKeys(sPath, sKey) 'parse keys
    
    Dim sData As String
    
    'open key
    If (RegOpenKeyEx(hKey, sKey, 0, KEY_READ, mainKey) = ERROR_SUCCESS) Then
        'query data
        If (RegQueryValueEx(mainKey, sName, 0, 0, sData, 1) = ERROR_SUCCESS) Then
            RegCloseKey mainKey 'close key
            ValueExists = True 'value exists
        Else
            ValueExists = False 'value don't exists
        End If
    Else
        ValueExists = False 'key don't exists
    End If
    
End Function
'----[ EnumValues ]----'
Public Function EnumValues(ByVal sPath As String, sName() As String, _
                sValue() As Variant, Optional OnlyType As rcRegType = -1) As Long
    Dim mainKey As Long, rName As String, Cnt As Long
    Dim rData As String, rType As Long, RetData As Long, RetVal As Long
    
    hKey = GetKeys(sPath, sKey) 'get key handles
    
    If RegOpenKey(hKey, sKey, mainKey) = ERROR_SUCCESS Then 'open key
        
        'reset data
        Cnt = 0
        rName = Space(255)
        rData = Space(255)
        RetVal = 255
        RetData = 255
        Erase sName
        Erase sValue
        
        'loop trought all values in this key
        While RegEnumValue(mainKey, Cnt, rName, RetVal, 0, _
                           rType, ByVal rData, RetData) <> ERROR_NO_MORE_ITEMS
                
            If (OnlyType = -1) Or (OnlyType = rType) Then 'is type we need?
                
                ReDim Preserve sName(EnumValues) As String 'englare arrays
                ReDim Preserve sValue(EnumValues) As Variant
                
                'name of value
                sName(EnumValues) = Trim$(Left$(rName, RetVal))
                
                'fill data
                If (rType = REG_BINARY) Then
                    If RetData > 0 Then 'BUG fixed here! ;)
                        sValue(EnumValues) = Trim$(BinToStr(Left$(rData, RetData)))
                    Else
                        sValue(EnumValues) = ""
                    End If
                ElseIf (rType = REG_DWORD) Then
                    sValue(EnumValues) = ReadDWORD(sPath, sName(EnumValues), 0)
                ElseIf (rType = REG_SZ) Then
                    sValue(EnumValues) = ReadString(sPath, sName(EnumValues), "")
                End If
                    
                EnumValues = EnumValues + 1 'incerment counter
            End If
                
                'prepare data for next item
                Cnt = Cnt + 1
                rName = Space(255)
                rData = Space(255)
                RetVal = 255
                RetData = 255
            
        Wend 'looping
        
        RegCloseKey hKey 'close key when we're done with it
    Else
        EnumValues = -1 'error
    End If
    
End Function
'----[ ExportToReg ]----'
Public Function ExportToReg(ByVal sPath As String, ByVal sRegFile As String, _
Optional IncludeSubkeys As Boolean = True, Optional ByVal Output As TextBox) As Long

    On Error GoTo errh
    Dim opened As Boolean, fn As Integer
    
    If (Dir(sRegFile) <> "") Then 'file allready exists
        ExportToReg = 0 'report error
        Exit Function
    End If
    
    fn = FreeFile 'get free number
    Open sRegFile For Output As #fn 'open it for writting
        opened = True 'set flag
        Print #fn, "REGEDIT4" & vbCrLf 'backward compitability
    Close #fn 'close it
    opened = False 'set flag
    
    
    If (generateReg(sPath, sRegFile, IncludeSubkeys, Output) = False) Then GoTo errh
    
    ExportToReg = 1 'successufuly generated .res file
    Exit Function
errh:
    On Error Resume Next
    If (opened = True) Then Close #fn
    ExportToReg = 0 'error
End Function
'----[ generateReg ]----'
Private Function generateReg(ByVal sPath As String, sRegFile As String, _
Optional IncludeSubkeys As Boolean = True, Optional Output As TextBox) As Boolean

    On Error GoTo errh
    Dim KeyName() As String, aName() As String, aValue() As Variant, X As Integer
    Dim u As Long, fn As Integer, tmp As String, opened As Boolean, l As Long
    Dim hasOutput As Boolean
    
    hasOutput = Not IsMissing(Output)
    
    'replace short with long key constants
    sPath = Replace(sPath, "HKCU", "HKEY_CURRENT_USER", , , 1)
    sPath = Replace(sPath, "HKLM", "HKEY_LOCAL_MACHINE", , , 1)
    sPath = Replace(sPath, "HKCR", "HKEY_CLASSES_ROOT", , , 1)
    sPath = Replace(sPath, "HKUS", "HKEY_USERS", , , 1)
    sPath = Replace(sPath, "HKPD", "HKEY_PERFORMANCE_DATA", , , 1)
    sPath = Replace(sPath, "HKDD", "HKEY_DYN_DATA", , , 1)
    sPath = Replace(sPath, "HKCC", "HKEY_CURRENT_CONFIG", , , 1)
    
    If (hasOutput = True) Then
        DoEvents
        Output.Text = sPath 'put current key name to textbox
    End If
        
    fn = FreeFile
    Open sRegFile For Append As #fn 'appending to file
        Print #fn, "[" & sPath & "]" 'print key name
        If (ReadString(sPath, "") <> vbNullChar) Then '(Default)
            Print #fn, "@=" & Chr$(34) & ReadString(sPath, "", "") & Chr$(34)
        End If
        
        u = EnumValues(sPath, aName, aValue, REG_SZ) - 1 'get all string values
        For l = 0 To u
            If (Len(aName(l)) > 0) Then
                Print #fn, Chr$(34) & aName(l) & Chr$(34) & "=" & _
                           Chr$(34) & aValue(l) & Chr$(34)
            End If
        Next
        u = EnumValues(sPath, aName, aValue, REG_BINARY) - 1 'get all binary values
        For l = 0 To u
            Print #fn, Chr$(34) & aName(l) & Chr$(34) & "=hex:" & _
                                                Replace(Trim$(aValue(l)), " ", ",")
        Next
        u = EnumValues(sPath, aName, aValue, REG_DWORD) - 1 'get all dword values
        For l = 0 To u
            tmp = "0x" & Right$("00000000" & Hex$(aValue(l)), 8)
            Print #fn, Chr$(34) & aName(l) & Chr$(34) & "=dword:" & tmp
        Next
        
        Print #fn, "" 'blank line
        
        On Error Resume Next 'close file
        Close #fn
        opened = False
        
        If (IncludeSubkeys = True) Then 'go trouhgt all subkeys if needed
            u = EnumKeys(sPath, KeyName) - 1
            For l = 0 To u
                If (generateReg(sPath & "\" & KeyName(l), sRegFile, _
                    IncludeSubkeys, Output) = False) Then GoTo errh
            Next
        End If
        
    Close #fn
    opened = False
    
    generateReg = True 'file written successufuly
    Exit Function
errh:
    On Error Resume Next
    If (opened = True) Then Close #fn
    generateReg = False 'error
End Function
'----[ ImportFromReg ]----'
Public Function ImportFromReg(ByVal sRegFile As String) As Long
    On Error GoTo noexists
    
    Dim Lines() As String, i As Long, sTemp As String, FileNum As Integer
    Dim Args() As String, k As Long, sLine As String, l As Long, tmp As String
    
    CreateKeyIfDoesntExists = True 'important!
    
    If (Dir(sRegFile) = "") Then 'file don't exists!
noexists:
        ImportFromReg = 0
        Exit Function
    End If

    FileNum = FreeFile
    Open sRegFile For Binary As #FileNum 'open file
        sTemp = Input(LOF(FileNum), #FileNum) 'and get all his contents
    Close #FileNum
    
    Lines = Split(sTemp, vbCrLf) 'split it in lines
    
    If (UCase$(Lines(0)) <> "REGEDIT4") Then
        ImportFromReg = 0 'reg file is NOT valid!
        Exit Function
    End If

    For i = 1 To UBound(Lines) 'for each line
        sLine = Replace(Trim$(Lines(i)), Chr$(34), vbNullString)

        If (Left$(sLine, 1) = "[") Then 'key
            sLine = Mid$(sLine, 2, Len(sLine) - 2) 'delete "[' and "]"
            
            If (Left$(sLine, 1) = "-") Then 'we need to kill this key
                sTemp = Mid$(sLine, 2, Len(sLine) - 1) 'remove "-"
                KillKey sTemp 'and delete it!
            Else
                For k = i + 1 To UBound(Lines)
                    sTemp = Trim$(Replace(Lines(k), Chr$(34), "")) 'remove quotes
                    
                    If (Left$(sTemp, 1) = "[") Then 'new key, return
                        i = k - 1
                        Exit For
                    End If
                    
                    If (sTemp = "") Or (InStr(1, sTemp, "=") < 1) Or _
                       (Left$(sTemp, 1) = ";") Then GoTo jump 'skip this line
                    
                    Args = Split(sTemp, "=") 'get arguments
                    
                    If (Trim$(Args(1)) = "-") Then 'delete value
                        KillValue sLine, Args(0)
                    Else 'adding value
                        If (LCase$(Left$(Args(1), 4)) = "hex:") Then 'binary
                            tmp = Replace(Mid$(Args(1), 5, Len(Args(1)) - 4), _
                                                                         ",", " ")
                            WriteBinary sLine, Args(0), tmp
                        ElseIf (LCase$(Left$(Args(1), 6)) = "dword:") Then 'DWORD
                            WriteDWORD sLine, Args(0), _
                                CLng(Val("&H" & Mid$(Args(1), 7, Len(Args(1)) - 6)))
                        Else 'string
                            WriteString sLine, Args(0), Args(1)
                            If (Args(0) = "@") And (Args(1) = "") Then _
                            KillValue sLine, "" '(value not set)
                        End If
                    End If
jump:
                Next
            End If
        End If
    Next
    
    ImportFromReg = 1 'success!
End Function
'----[ StrToBin ]----'
Public Function StrToBin(sBin As String) As String
    Dim two() As String, q As Integer
    Dim bs As String, w As Integer
    
    sBin = Trim$(Replace(sBin, " ", vbNullString)) 'remove spaces
    
    If Len(sBin) = 0 Then Exit Function
    
    ReDim two(1 To Len(sBin)) As String
    
    w = 0
    For q = 1 To Len(sBin) Step 2 'two by two
        w = w + 1
        bs = Mid$(sBin, q, 2)
        If bs = "00" Then bs = vbNullChar
        two(w) = bs
    Next

    For q = 1 To UBound(two) / 2
        If two(q) = vbNullChar Then
            StrToBin = StrToBin & vbNullChar
        Else
            StrToBin = StrToBin & Chr$(Val("&H" & two(q)))
        End If
    Next
End Function
'----[ BinToStr ]----'
Public Function BinToStr(sStr As String) As String
    'will convert eg "¾>ÿ«" into "BE 3E FF AB" - used for binary data
    Dim bs As String, ret As String, q As Integer, tStr As String
    
    If Len(sStr) = 0 Then GoTo zero_length
    
    ret = vbNullString
    For q = 1 To Len(sStr)
        bs = Mid$(sStr, q, 1)
        If bs = vbNullChar Then tStr = "00" Else tStr = CStr(Hex(Asc(bs)))
        If (Len(tStr) = 1) Then tStr = tStr & "0"
        ret = ret & " " & tStr
    Next
zero_length:
    BinToStr = ret
End Function
'----[ isBinValid ]----'
Public Function isBinValid(ByVal sBin As String) As Boolean
    Dim z As Long
    
    sBin = Trim$(UCase$(Replace(sBin, " ", vbNullString)))
    
    If Len(sBin) = 0 Then GoTo zero_length
    
    For z = 1 To Len(sBin)
        If InStr(1, Mid$(sBin, z, 1), "0123456789ABCDEF ", 1) < 1 Then
zero_length:
           isBinValid = False
           Exit Function
        End If
    Next
    isBinValid = True
    
End Function

Private Sub Class_Initialize()
    CreateKeyIfDoesntExists = True 'default
End Sub


'p.s. sorry for all grammar and spelling errors, my english sux ;o)


