Attribute VB_Name = "Security"
Option Explicit

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003
Public Const ERROR_SUCCESS = 0&
Public Const REG_OPTION_NON_VOLATILE = &O0
Public Const KEY_ALL_CLASSES As Long = &HF0063
Public Const KEY_ALL_ACCESS = &H3F
Public Const REG_SZ As Long = 1
Public Const REG_DWORD = 4

Public Const VisPath = "SOFTWARE\JAWOOTEK\Parking"

'Type SYSTEMTIME
'        wYear As Integer
'        wMonth As Integer
'        wDayOfWeek As Integer
'        wDay As Integer
'        wHour As Integer
'        wMinute As Integer
'        wSecond As Integer
'        wMilliseconds As Integer
'End Type
'Declare Function SetLocalTime Lib "KERNEL32" (lpSystemTime As SYSTEMTIME) As Long

Public Sub Time_Sync()
    
    Dim qry As String
    Dim rs As ADODB.Recordset

On Error GoTo Err_P

    Set rs = New ADODB.Recordset
    qry = "Select date_format(now()," & Chr(34) & "%Y%m%d%H%i%S" & Chr(34) & ");"
    rs.Open qry, adoConn
    
    If Not (rs.EOF) Then
        'Debug.Print rs(0)
        If (rs(0) <> Format(Now, "yyyymmddhhnn")) Then
            Call Set_Time(rs(0))
            Call DataLogger("DB Time Sync. Success..!!")
        End If
    End If
    Set rs = Nothing

Exit Sub

Err_P:
    Call DataLogger("TimeSync Proc Error")

End Sub


Public Sub Set_Time(time_str As String)
    
    Dim Sys_Time As SYSTEMTIME
    Dim tmp As Long
    Dim YYYY As Integer
    Dim MM As Integer
    Dim DD As Integer
    Dim HH As Integer
    Dim NN As Integer
    Dim SS As Integer
    
On Error GoTo Err_P
    
    'yyyymmddhhnnss
    YYYY = Mid(time_str, 1, 4)
    MM = Mid(time_str, 5, 2)
    DD = Mid(time_str, 7, 2)
    HH = Mid(time_str, 9, 2)
    NN = Mid(time_str, 11, 2)
    SS = Mid(time_str, 13, 2)
    Sys_Time.wYear = YYYY
    Sys_Time.wMonth = MM
    Sys_Time.wDayOfWeek = 0
    Sys_Time.wDay = DD
    Sys_Time.wHour = HH
    Sys_Time.wMinute = NN
    Sys_Time.wSecond = SS
    Sys_Time.wMilliseconds = 0
    tmp = SetLocalTime(Sys_Time)

Exit Sub

Err_P:

End Sub


Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, vValue, Len(vValue))
End Function

Public Function GetRegValue(TopKey As Long, SubKey As String, ValueTitle As String) As String
    Dim lRetVal As Long
    Dim Buffer As String * 128 '버퍼
    Dim lBufferSize As Long '버퍼크기
    Dim lSubKey As Long
    Dim dType As Long
    Dim i As Integer
    Dim xx As String
    Dim tt As Double

    lBufferSize = 64

    lRetVal = RegOpenKeyEx(TopKey, SubKey, 0, KEY_ALL_ACCESS, lSubKey)
    If lRetVal <> ERROR_SUCCESS Then
        GetRegValue = ""
        RegCloseKey lSubKey
        Exit Function
    End If

   lRetVal = RegQueryValueEx(lSubKey, ValueTitle, 0, dType, ByVal Buffer, lBufferSize)
   If lRetVal = ERROR_SUCCESS Then
        If dType = 4 Then '레지스트리의 값 타입이 Double Word 형이면
            For i = 1 To 4
                xx = Mid(Buffer, i, 1)
                tt = tt + Asc(xx) * 256 ^ (i - 1)
            Next i
                GetRegValue = Trim(tt)
        Else
            xx = ""
            For i = 1 To Len(Buffer)
                If Mid(Buffer, i, 1) = Chr(0) Then Exit For
                xx = xx + Mid(Buffer, i, 1)
            Next i
            GetRegValue = xx
        End If
    Else
        GetRegValue = "" '(원본)레지스트리 키 이름이 없으면 에러가 난다
        'GetRegValue = 0 '(내가고친것)에러나는 이유는? 모르겠음
        RegCloseKey lSubKey
        Exit Function
    End If
    RegCloseKey lSubKey '열려진 레지스트리 키를 닫는다
End Function

Public Function SetRegValue(TopKey As Long, SubKey As String, ValueTitle As String, value As String, dType As Long) As String
    Dim lSubKey As Long
    Dim lRetVal As Long
    Dim Ivalue As Long

    lRetVal = RegCreateKey(TopKey, SubKey, lSubKey)

    If lRetVal <> ERROR_SUCCESS Then
        SetRegValue = ""
        RegCloseKey lSubKey
        Exit Function
    End If

    If dType = 4 Then '레지스트리의 값 타입이 Double Word 형이면
        Ivalue = Val(value)
        lRetVal = RegSetValueEx(lSubKey, ValueTitle, 0, REG_DWORD, Ivalue, 4)
        If lRetVal = ERROR_SUCCESS Then
            SetRegValue = value
        Else
            SetRegValue = ""
        End If
    Else
        If value = "" Then value = " "
        'lRetVal = RegSetValueEx(lSubKey, ValueTitle, 0, REG_SZ, ByVal Value, Len(Value) + 1) 리부팅 시간의 마지막 초가 짤려서 길이를 임으로 늘렸다. 문제?
        lRetVal = RegSetValueEx(lSubKey, ValueTitle, 0, REG_SZ, ByVal value, Len(value) + 2)
        If lRetVal = ERROR_SUCCESS Then
            SetRegValue = value
        Else
            SetRegValue = ""
        End If
    End If
End Function

Public Function GetDriveSerialNumber(Optional ByVal DriveLetter As String) As String


    Dim fso As Object, Drv As Object
    Dim DriveSerial As String
    Dim strTemp As String
    
    'Create a FileSystemObject object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Assign the current drive letter if not specified
    If DriveLetter <> "" Then
        Set Drv = fso.GetDrive(DriveLetter)
    Else
        Set Drv = fso.GetDrive(fso.GetDriveName(App.Path))
    End If

    With Drv
        If .IsReady Then
            DriveSerial = Abs(.SerialNumber)
        Else    '"Drive Not Ready!"
            DriveSerial = -1
        End If
    End With
    
    'Clean up
    Set Drv = Nothing
    Set fso = Nothing
    strTemp = Hex(DriveSerial)
    GetDriveSerialNumber = Left(strTemp, 4) & Right(strTemp, 4)
    
End Function

'XOR 알고리즘
Public Function Encrypt(ByRef Original As String) As String
    If LenB(Original) = 0 Or Original = Null Then Exit Function
    
    Dim buf() As Byte
    Dim KEY As Byte
    KEY = 11
    
    buf() = StrConv(Original, vbFromUnicode)
    
    Dim i As Long
    
    For i = 0 To UBound(buf)
        Encrypt = Encrypt & Right$("0" & Hex$(buf(i) Xor KEY), 2)
    Next
End Function

Public Function Decrypt(ByRef Crypted As String) As String
    If LenB(Crypted) = 0 Or Crypted = Null Then Exit Function
    
    Dim i As Long
    Dim KEY As Byte
    KEY = 11
        
    If Crypted = " " Then
        Exit Function
    End If
        
    For i = 1 To Len(Crypted) Step 2
        Decrypt = Decrypt & ChrB$(CByte("&H" & Mid$(Crypted, i, 2)) Xor KEY)
    Next
    
    Decrypt = StrConv(Decrypt, vbUnicode)
End Function

Public Sub HomeLogger(LogStr As String)
    Dim intFileNum As Integer
    
    intFileNum = FreeFile()
    Open Doc_Path_Name$ & "HomeLog_" & Format(Now, "yyyy-mm-dd") & ".txt" For Append As #intFileNum
    Print #intFileNum, "HomeLog_" & Format(Now, "yyyy-mm-dd hh:nn:ss ") & "    " & LogStr
    Close #intFileNum
    
    Select Case HomeNetMode
        Case 1
            FrmHyundae.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 2
            FrmGyeyoung.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 3
            FrmEZVille.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 4
            FrmKocom.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 5
            FrmCommax.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 6
            FrmIControls.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 7
            FrmKDOne.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 8
            FrmLGElectron.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 9
            FrmGyeyoungTCP.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 10
            FrmHyundae_Linux.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 11
            FrmMaxuracy.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 12
            FrmHomeclever.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 13
            frmGyeyoungRS232.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
    End Select
    
    
    'Textbox 가로 스크롤바
    If (LenH(LogStr) > Glo_MaxLogLen) Then
        
        Dim MyForm As Form
        Dim nMaxScale As Long
        
        Select Case HomeNetMode
            Case 1
                Set MyForm = FrmHyundae
            Case 2
                Set MyForm = FrmGyeyoung
            Case 3
                Set MyForm = FrmEZVille
            Case 4
                Set MyForm = FrmKocom
            Case 5
                Set MyForm = FrmCommax
            Case 6
                Set MyForm = FrmIControls
            Case 7
                Set MyForm = FrmKDOne
            Case 8
                Set MyForm = FrmLGElectron
            Case 9
                Set MyForm = FrmGyeyoungTCP
            Case 10
                Set MyForm = FrmHyundae_Linux
            Case 11
                Set MyForm = FrmMaxuracy
            Case 12
                Set MyForm = FrmHomeclever
            Case 13
                Set MyForm = frmGyeyoungRS232
        End Select
    
        Glo_MaxLogLen = LenH(LogStr)
        nMaxScale = MyForm.TextWidth(MyForm.List1.List(0))
        SendMessageByNum MyForm.List1.hwnd, LB_SETHORIZONTALEXTENT, nMaxScale / 15 + 250, 0
        
        Set MyForm = Nothing
    End If
    

End Sub

Public Sub DataLogger(LogStr As String)
'Public Sub Err_doc(Err_str As String)
    Dim intFileNum As Integer
    
    intFileNum = FreeFile()
    Open Doc_Path_Name$ & Format(Now, "yyyy-mm-dd") & ".txt" For Append As #intFileNum
    Print #intFileNum, Format(Now, "yyyy-mm-dd hh:nn:ss ") & "    " & LogStr
    Close #intFileNum
    
    Select Case HomeNetMode
        Case 1
            FrmHyundae.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 1
            FrmHyundae.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 2
            FrmGyeyoung.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 3
            FrmEZVille.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 4
            FrmKocom.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 5
            FrmCommax.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 6
            FrmIControls.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 7
            FrmKDOne.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 8
            FrmLGElectron.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 9
            FrmGyeyoungTCP.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 10
            FrmHyundae_Linux.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 11
            FrmMaxuracy.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 12
            FrmHomeclever.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
        Case 13
            frmGyeyoungRS232.List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0
    End Select


End Sub
