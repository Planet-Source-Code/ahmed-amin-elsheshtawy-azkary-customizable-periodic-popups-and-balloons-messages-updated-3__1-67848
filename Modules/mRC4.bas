Attribute VB_Name = "mRC4"
'**************************************
' Name: (Update) RC4 Stream Cipher (with
'     file handling )
' Description:This code offers you a str
'     ong encryption with RC4. I've tested it
'     a lot and it's the right implementation
'     of the RC4 cipher.
'You can use this code in your commercia
'     l code because it's not patented!
'I know there is another code that deals


'     with RC4 but my code has nothing to do w
    '     ith this code!
    'More infos: sci.crypt
' By: Sebastian
'
'
' Inputs:Create the form and simply sele
'     ct a file to en(de)crypt!
'Notice that you use the same function f
'     or encryption and decryption
'
' Returns:After you press the Button you
'     should get the en(de)crypted file!
'
'Assumes:'Assumes:Create a form with:
'
'txtpwd (txtbox)
'txtSave (txtbox)
'txtPattern (Combobox)
'filList (FileListBox)
'DirList (DirListBox)
'drvList (DrvlistBox)
'Command1 (Command Button ; Caption=Encr
'     ypt)
'Command2 (Command Button ; Caption=Decr
'     ypt)
'
'Side Effects:If you encrypt different t
'     extes with the same password, someone co
'     uld be able to decrypt your code. (This
'     is quiet normal for a stream cipher!)
'IF YOU ENCRYPT LARGE FILES PLEASE USE THE EnDeCryptSingle ROUTINE INSTEAD OF THE EnDeCrypt ROUTINE OR SPLIT THE INPUT IN SMALLER PIECES!
'This code is copyrighted and has limite
'     d warranties.
'Please see http://www.Planet-Source-Cod
'     e.com/xq/ASP/txtCodeId.1736/lngWId.1/qx/
'     vb/scripts/ShowCode.htm
'for details.
'**************************************

Option Explicit
Private s(0 To 255) As Integer  'S-Box
Private kep(0 To 255) As Integer
Private i As Integer
Private j As Integer

Public Sub RC4Ini(Pwd As String)
    Dim temp As Integer, a As Integer, b As Integer
    'Save Password in Byte-Array
    b = 0

    For a = 0 To 255
        b = b + 1
        If b > Len(Pwd) Then
            b = 1
        End If
        kep(a) = Asc(Mid$(Pwd, b, 1))
    Next a
    'INI S-Box

    For a = 0 To 255
        s(a) = a
    Next a
    b = 0

    For a = 0 To 255
        b = (b + s(a) + kep(a)) Mod 256
        ' Swap( S(i),S(j) )
        temp = s(a)
        s(a) = s(b)
        s(b) = temp
    Next a
End Sub

'Only use this routine for short texts
Public Function EnDeCrypt(plaintxt As Variant) As Variant

    Dim temp As Integer, a As Long, i As Integer
    Dim j As Integer, k As Integer
    Dim cipherby As Byte, cipher As Variant

    For a = 1 To Len(plaintxt)
        i = (i + 1) Mod 256
        j = (j + s(i)) Mod 256
        ' Swap( S(i),S(j) )
        temp = s(i)
        s(i) = s(j)
        s(j) = temp
        'Generate Keybyte k
        k = s((s(i) + s(j)) Mod 256)
        'Plaintextbyte xor Keybyte
        cipherby = Asc(Mid$(plaintxt, a, 1)) Xor k
        cipher = cipher & Chr(cipherby)
    Next a
    
    EnDeCrypt = cipher
    
End Function

'Use this routine for really huge files
Public Function EnDeCryptSingle(plainbyte As Byte) As Byte
    
    Dim temp As Integer, k As Integer
    Dim cipherby As Byte
    
    i = (i + 1) Mod 256
    j = (j + s(i)) Mod 256
    ' Swap( S(i),S(j) )
    temp = s(i)
    s(i) = s(j)
    s(j) = temp
    'Generate Keybyte k
    k = s((s(i) + s(j)) Mod 256)
    'Plaintextbyte xor Keybyte
    cipherby = plainbyte Xor k
    EnDeCryptSingle = cipherby
    
End Function

'Private Sub Test()
'
'    Dim Text As String
'    RC4Ini "ahmed"
'    Text = EnDeCrypt("Hello")
'    Debug.Print Text & vbCrLf
'
'    RC4Ini "ahmed"
'    Text = EnDeCrypt(Text)
'    Debug.Print "Decrypted: " & Text & vbCrLf
'
'End Sub
