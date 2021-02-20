Attribute VB_Name = "Module3"
' $Id: UtfTests.bas $

' Converting a VBA string to an array of bytes in UTF-8 encoding

' $Date: 2015-06-30 10:05Z $
' $Author: David Ireland $

' Copyright (C) 2015 DI Management Services Pty Limited
' <http://www.di-mgt.com.au> <http://www.cryptosys.net>

Option Explicit
Option Base 0

''' Extract a set of VBA "Unicode" strings from Excel sheet, encode in UTF-8 and display details
Public Sub ShowStuff()
    Dim strData As String
    
    ' Plain ASCII
    ' "abc123"
    ' U+0061, U+0062, U+0063, U+0031, U+0032, U+0033
    ' EXCEL: Get value from cell A1
    strData = Worksheets("Sheet1").Cells(1, 1)
    Debug.Print vbCrLf & Worksheets("Sheet1").Cells(1, 2)
    ProcessString (strData)

    ' Spanish
    ' LATIN SMALL LETTER[s] [AEIO] WITH ACUTE and SMALL LETTER N WITH TILDE
    ' U+00E1, U+00E9, U+00ED, U+00F3, U+00F1
    ' EXCEL: Get value from cell A3
    strData = Worksheets("Sheet1").Cells(3, 1)
    Debug.Print vbCrLf & Worksheets("Sheet1").Cells(3, 2)
    ProcessString (strData)

    ' Japanese
    ' "Hello" in Hiragana characters is KO-N-NI-TI-HA (Kon'nichiwa)
    ' U+3053 (hiragana letter ko), U+3093 (hiragana letter n),
    ' U+306B (hiragana letter ni), U+3061 (hiragana letter ti),
    ' and U+306F (hiragana letter ha)
    ' EXCEL: Get value from cell A5
    strData = Worksheets("Sheet1").Cells(5, 1)
    Debug.Print vbCrLf & Worksheets("Sheet1").Cells(5, 2)
    ProcessString (strData)

    ' Chinese
    ' CN=ben (U+672C), C= zhong guo (U+4E2D, U+570B), OU=zong ju (U+7E3D, U+5C40)
    ' EXCEL: Get value from cell A7
    strData = Worksheets("Sheet1").Cells(7, 1)
    Debug.Print vbCrLf & Worksheets("Sheet1").Cells(7, 2)
    ProcessString (strData)

    ' Hebrew
    ' "abc" U+0061, U+0062, U+0063
    ' SPACE U+0020
    ' [NB right-to-left order]
    ' U+05DB HEBREW LETTER KAF
    ' U+05E9 HEBREW LETTER SHIN
    ' U+05E8 HEBREW LETTER RESH
    ' SPACE "f123" U+0066 U+0031 U+0032 U+0033
    ' EXCEL: Get value from cell A9
    strData = Worksheets("Sheet1").Cells(9, 1)
    Debug.Print vbCrLf & Worksheets("Sheet1").Cells(9, 2)
    ProcessString (strData)
    
End Sub

Public Function ProcessString(strData As String)
    Dim abData() As Byte
    Dim strOutput As String
    
    Debug.Print strData ' This should show "?" for non-ANSI characters
    abData = Utf8BytesFromString(strData)
    Debug.Print bv_HexFromBytesSp(abData)
    Debug.Print "Strlen=" & Len(strData) & " chars; utf8len=" & UBound(abData) + 1 & " bytes"

End Function

''' Returns hex-encoded string from array of bytes (with spaces)
''' E.g. aBytes(&HFE, &HDC, &H80) will return "FE DC 80"
Public Function bv_HexFromBytesSp(aBytes() As Byte) As String
    Dim i As Long

    If Not IsArray(aBytes) Then
        Exit Function
    End If
    
    For i = LBound(aBytes) To UBound(aBytes)
        If (i > 0) Then bv_HexFromBytesSp = bv_HexFromBytesSp & " "
        If aBytes(i) < 16 Then
            bv_HexFromBytesSp = bv_HexFromBytesSp & "0" & Hex(aBytes(i))
        Else
            bv_HexFromBytesSp = bv_HexFromBytesSp & Hex(aBytes(i))
        End If
    Next
    
End Function



