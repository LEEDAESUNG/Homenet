Attribute VB_Name = "Module2"

' basUtf8FromString

' Written by David Ireland DI Management Services Pty Limited 2015
' <http://www.di-mgt.com.au> <http://www.cryptosys.net>

Option Explicit

''' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long
    
' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

''' Return byte array with VBA "Unicode" string encoded in UTF-8
Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, vbNull, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function



Public Function StringToHex(ByVal StrToHex As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim I         As Long
    For I = 1 To Len(StrToHex)
        strTemp = Hex$(Asc(Mid$(StrToHex, I, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & Space$(1) & strTemp
    Next I
    StringToHex = strReturn
End Function


Public Function StringToHex2(ByVal StrToHex As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim I         As Long
    For I = 1 To LenH(StrToHex)
        strTemp = Hex$(Asc(Mid$(StrToHex, I, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & Space$(1) & strTemp
    Next I
    StringToHex2 = strReturn
End Function


Public Function HexToString(ByVal HexToStr As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim I         As Long
    For I = 1 To Len(HexToStr) Step 3
        strTemp = Chr$(Val("&H" & Mid$(HexToStr, I, 2)))
        strReturn = strReturn & strTemp
    Next I
    HexToString = strReturn
End Function



Public Function ByteArrayToHex(ByRef ByteArray() As Byte) As String
    Dim lb As Long, ub As Long
    Dim l As Long, strRet As String
    Dim lonRetLen As Long, lonPos As Long
    Dim strHex As String, lonLenHex As Long
    
    lb = LBound(ByteArray)
    ub = UBound(ByteArray)
    lonRetLen = ((ub - lb) + 1) * 3
    strRet = Space$(lonRetLen)
    lonPos = 1
    
    For l = lb To ub
        strHex = Hex$(ByteArray(l))
        If Len(strHex) = 1 Then strHex = "0" & strHex
        If l <> ub Then
            Mid$(strRet, lonPos, 3) = strHex & " "
            lonPos = lonPos + 3
        Else
            Mid$(strRet, lonPos, 3) = strHex
        End If
    Next l
    
    ByteArrayToHex = strRet
End Function

