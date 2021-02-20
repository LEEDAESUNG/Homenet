Attribute VB_Name = "Definition"

Option Explicit

Public Glo_MaxLogLen As Long

' 세대통보 패킷 데이터 저장 전역변수(시리얼통신용)
Public GbDATA As Boolean
Public GsData As String
Public Glo_Event As Boolean
Public Glo_bData(1 To 14) As Byte
Public Glo_chk As Byte




Public HC_Tcp_Header() As Byte
Public GY_Tcp_Header() As Byte

Public KD_Header() As Byte

Public IniFileName$
Public AdoHome_Str As String
Public rs As ADODB.Recordset
Public Homers As ADODB.Recordset

Public Report_Path_Name$
Public Doc_Path_Name$

Public Local_IP As String

Public PassWord As String
Public kyo_str(33) As String * 30

Public Glo_MsgRet As Boolean
Public Server_IP As String
Public HostPort As Long

Public HomeNetMode As Integer
Public HomeNet_IP As String
Public HomeNet_Port As Long
Public HomeNet_ComPort As Long
Public HomeNet_ID As String
Public HomeNet_PW As String

Public ezVille_Dong As String
Public ezVille_Ho As String

Public HomeNet_Dong As String * 4
Public HomeNet_Ho As String * 4
Public HomeNet_CarNo As String * 16

'Localhost Config
Public LocalHostIP As String
Public Socket_Data As String

'''' WinApi function that maps a UTF-16 (wide character) string to a new character string
'Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
'    ByVal CodePage As Long, _
'    ByVal dwFlags As Long, _
'    ByVal lpWideCharStr As Long, _
'    ByVal cchWideChar As Long, _
'    ByVal lpMultiByteStr As Long, _
'    ByVal cbMultiByte As Long, _
'    ByVal lpDefaultChar As Long, _
'    ByVal lpUsedDefaultChar As Long) As Long
'
'Private Const CP_UTF8 = 65001
'
'''' Return byte array with VBA "Unicode" string encoded in UTF-8
'Public Function Utf8BytesFromString(strInput As String) As Byte()
'    Dim nBytes As Long
'    Dim abBuffer() As Byte
'    ' Get length in bytes *including* terminating null
'    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, vbNull, 0&, 0&, 0&)
'    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
'    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
'    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
'    Utf8BytesFromString = abBuffer
'End Function

'전역변수 초기화
Public Sub CFG_Init()

    '1.현대통신 2.서울통신(MySQL)
    HomeNetMode = Val(Get_Ini("System Config", "HomeNetMode", "0"))
    HostPort = Val(Get_Ini("System Config", "HostPort", "18497"))
    
    Select Case HomeNetMode
        Case 1
            HomeNet_IP = Get_Ini("Hyundae", "HomeNet_IP", "127.0.0.1")
            HomeNet_Port = Val(Get_Ini("Hyundae", "HomeNet_Port", "4000"))
        
        Case 2
            AdoHome_Str = Get_Ini("Seoul_DB", "HomeNet_Str", "")
    
        Case 3
            HomeNet_IP = Get_Ini("ezVille", "HomeNet_IP", "127.0.0.1")
            HomeNet_Port = Val(Get_Ini("ezVille", "HomeNet_Port", "4000"))
            ezVille_Dong = Get_Ini("ezVille", "ezVille_Dong", "100")
            ezVille_Ho = Get_Ini("ezVille", "ezVille_Ho", "100")
            
        Case 4
            HomeNet_IP = Get_Ini("Kocom", "HomeNet_IP", "127.0.0.1")
            HomeNet_Port = Val(Get_Ini("Kocom", "HomeNet_Port", "4000"))
            HomeNet_ID = Get_Ini("Kocom", "HomeNet_ID", "parkone")
            HomeNet_PW = Get_Ini("Kocom", "HomeNet_PW", "parkone1234")
            
        Case 5
            HomeNet_IP = Get_Ini("Commax", "HomeNet_IP", "127.0.0.1")
            HomeNet_Port = Val(Get_Ini("Commax", "HomeNet_Port", "4000"))
            
        Case 6
            Local_IP = Get_Ini("IControls", "Local_IP", "127.0.0.1")
            HomeNet_IP = Get_Ini("IControls", "HomeNet_IP", "127.0.0.1")
            HomeNet_Port = Val(Get_Ini("IControls", "HomeNet_Port", "4000"))
        
        Case 7
            HomeNet_IP = Get_Ini("KD One", "HomeNet_IP", "127.0.0.1")
            HomeNet_Port = Val(Get_Ini("KD One", "HomeNet_Port", "4000"))
            
        Case 8
            HomeNet_IP = Get_Ini("LGElectron", "HomeNet_IP", "127.0.0.1")
            HomeNet_Port = Val(Get_Ini("LGElectron", "HomeNet_Port", "4000"))
            
        Case 9
            HomeNet_IP = Get_Ini("Seoul_TCP", "HomeNet_IP", "127.0.0.1")
            HomeNet_Port = Val(Get_Ini("Seoul_TCP", "HomeNet_Port", "4000"))
        
        Case 10
            HomeNet_IP = Get_Ini("Hyundae", "HomeNet_IP", "127.0.0.1")
            HomeNet_Port = Val(Get_Ini("Hyundae", "HomeNet_Port", "4000"))
            
        Case 11
            HomeNet_IP = Get_Ini("Maxuracy", "HomeNet_IP", "127.0.0.1")
            HomeNet_Port = Val(Get_Ini("Maxuracy", "HomeNet_Port", "4000"))
            
        Case 12
            HomeNet_IP = Get_Ini("Homeclever", "HomeNet_IP", "127.0.0.1")
            HomeNet_Port = Val(Get_Ini("Homeclever", "HomeNet_Port", "4000"))
        
        Case 13
            HomeNet_IP = ""
            HomeNet_Port = 0
            HomeNet_ComPort = Val(Get_Ini("Seoul_SERIAL", "HomeNet_ComPort", "1"))

        Case Else
    
    End Select

    

End Sub














































