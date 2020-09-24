Attribute VB_Name = "mParseURL"
Option Explicit

Function ParseURL(strURL, strPart)

'descr: parses a portion of a url
'strURL:    the URL to parse
'strPart:   the part to get. Allowed values are:
'   protocol, server, domain, path, file, hash, query

    Dim arrTemp
    Dim strTemp
    Dim nPos

    On Error Resume Next
    
    Select Case strPart
    
        Case "protocol"
            'return the protocol, eg. http://, ftp://
            nPos = InStr(strURL, ":") + 1
            Do Until (Mid(strURL, nPos, 1) <> "/") And (Mid(strURL, nPos, 1) <> "\")
                nPos = nPos + 1
            Loop
            ParseURL = Left(strURL, nPos - 1)
            
        Case "server"
            'return the server, eg. www.microsoft.com
            strTemp = ParseURL(strURL, "protocol")
            strURL = Right(strURL, Len(strURL) - Len(strTemp))
            If InStr(strURL, "/") Then
                strTemp = Left(strURL, InStr(strURL, "/") - 1)
            ElseIf InStr(strURL, "\") Then
                strTemp = Left(strURL, InStr(strURL, "\") - 1)
            End If
            If InStr(strTemp, "@") Then
            'remove user/password combo, return only the server
                ParseURL = Right(strTemp, Len(strTemp) - InStr(strTemp, "@"))
            Else
                ParseURL = strTemp
            End If
            
        Case "domain"
        'return only the domain, eg. amazon.com, wa.gov, etc
            strTemp = ParseURL(strURL, "server")
            arrTemp = Split(strTemp, ".")
            ParseURL = arrTemp(UBound(arrTemp) - 1) & "." & arrTemp(UBound(arrTemp))
        
        Case "path"
        'return the path
            If InStr(strURL, "#") Then strURL = Left(strURL, InStr(strURL, "#") - 1)
            If InStr(strURL, "?") Then strURL = Left(strURL, InStr(strURL, "?") - 1)
            If InStrRev(strURL, "/") > InStrRev(strURL, "\") Then
                ParseURL = Left(strURL, InStrRev(strURL, "/"))
            ElseIf InStrRev(strURL, "\") > InStrRev(strURL, "/") Then
                ParseURL = Left(strURL, InStrRev(strURL, "\"))
            End If
            
        Case "file"
        'return the filename only
            If InStr(strURL, "#") Then strURL = Left(strURL, InStr(strURL, "#") - 1)
            If InStr(strURL, "?") Then strURL = Left(strURL, InStr(strURL, "?") - 1)
            If InStrRev(strURL, "/") > InStrRev(strURL, "\") Then
                ParseURL = Right(strURL, Len(strURL) - InStrRev(strURL, "/"))
            ElseIf InStrRev(strURL, "\") > InStrRev(strURL, "/") Then
                ParseURL = Right(strURL, Len(strURL) - InStrRev(strURL, "\"))
            End If
            
        Case "hash"
        'return the bookmark (hash) without the hash mark
            If InStr(strURL, "#") Then
                arrTemp = Split(strURL, "#")
                strTemp = arrTemp(UBound(arrTemp))
                If InStr(strTemp, "?") Then
                    ParseURL = Left(strTemp, InStr(strTemp, "?") - 1)
                Else
                    ParseURL = strTemp
                End If
            Else
                ParseURL = ""
            End If
            
        Case "query"
        'return the query string without the question mark
            If InStr(strURL, "?") Then
                arrTemp = Split(strURL, "?")
                strTemp = arrTemp(UBound(arrTemp))
                If InStr(strTemp, "#") Then
                    ParseURL = Left(strTemp, InStr(strTemp, "#") - 1)
                Else
                    ParseURL = strTemp
                End If
            Else
                ParseURL = ""
            End If
            
    End Select
    
    If Err.Number <> 0 Then ParseURL = ""
        
End Function

