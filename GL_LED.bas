Attribute VB_Name = "GL_LED"
Option Explicit

Public Led_Show As Byte
Public Led_Speed As Byte
Public Led_StopTime As Byte
Public Led_Repeat As Byte
Public Led_up_color As Byte
Public Led_down_color As Byte

Public Nomal_Show As Byte
Public Nomal_Speed As Byte
Public Nomal_StopTime As Byte
Public Nomal_Up_color As Byte
Public Nomal_Down_color As Byte

Public Nomaltxt_Up As String
Public Nomaltxt_Down As String

Public Up_Color As Byte
Public Down_Color As Byte

Public Led_OutF_In As Byte
Public Led_OutF_Out As Byte

'Public M As String
Public C As String
Public ShowEffect As Byte
Public ShowSpeed As Byte
Public ShowTime As Byte
Public Repeat As Byte

Public In_Led_F As Boolean
Public Out_Led_F As Boolean

Public GloDisp_BData() As Byte
Public GlO_TcpDataGate As String


Public Sub Dis_Refresh(ByVal GateNo As Integer)
    Dim i As Integer
    Dim Qry As String
    Dim Ret As Integer

On Error GoTo Err_Proc

    Ret = DataBaseOpen(adoConn)
    If (Ret = 1) Then
    Else
        Call Err_doc("    [LPRIn_Proc] DB Connection Fail...!!! Disp_Refresh")
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    Qry = "SELECT * From tb_lpr WHERE GateNo = '" & GateNo & "'"
    rs.Open Qry, adoConn

    If Not (rs.EOF) Then
        Call GL_Nomal(Trim(rs!Dis1), Trim(rs!Dis2), 129, 70, 0, 2, 1, 0)
    End If
    Set rs = Nothing

    If (Ret = 1) Then
        Call DataBaseClose(adoConn)
    End If
    
Exit Sub

Err_Proc:
    Call Err_doc("    [DisRefresh_Proc] " & Err.Description)
    If (Ret = 1) Then
        Call DataBaseClose(adoConn)
    End If

End Sub


Public Sub GL_PowerOn(Port As Byte)
'ReDim a(9) As Byte
'
'a(0) = &H10     'DLE
'a(1) = &H2      'STX
'a(2) = &H0      'DST
'
'a(3) = &H0      'LEN
'a(4) = &H3
'
'a(5) = &H41     'CMD
'
'a(6) = &H1      'DATA : ON
'a(7) = &H0      '수직분할 화면에서 소화면 전원제어 0:OFF 1: ON
'
'a(8) = &H10
'a(9) = &H3
'
'With Jung
'If (.MSCommDisp(Port).PortOpen) Then
'    .MSCommDisp(Port).Output = a
'End If
'End With
End Sub

Public Sub GL_PowerOff(Port As Byte)
'ReDim a(9) As Byte
'
'a(0) = &H10     'DLE
'a(1) = &H2      'STX
'a(2) = &H0      'DST
'
'a(3) = &H0      'LEN
'a(4) = &H3
'
'a(5) = &H41     'CMD
'
'a(6) = &H0      'DATA : OFF
'a(7) = &H0      '수직분할 화면에서 소화면 전원제어 0:OFF 1: ON
'
'a(8) = &H10
'a(9) = &H3
'
'With Jung
'If (.MSCommDisp(Port).PortOpen) Then
'    .MSCommDisp(Port).Output = a
'End If
'End With
End Sub

Public Sub GL_Start(Port As Byte)
ReDim a(9) As Byte

a(0) = &H10     'DLE
a(1) = &H2      'STX
a(2) = &H0      'DST

a(3) = &H0      'LEN
a(4) = &H3

a(5) = &H45     'CMD

a(6) = &H1      'DATA1 : 0 :현상태 유지 1: Start 2:Stop BMP Display
a(7) = &H1      'DATA2 : 0 :현상태 유지 1: Start 2:Stop 긴급문구 Display

a(8) = &H10     '
a(9) = &H3

'With Jung
'If (.MSCommDisp(Port).PortOpen) Then
'    .MSCommDisp(Port).Output = a
'End If
'End With
End Sub

Public Sub GL_Stop(Port As Byte)
ReDim a(9) As Byte

a(0) = &H10     'DLE
a(1) = &H2      'STX
a(2) = &H0      'DST

a(3) = &H0      'LEN
a(4) = &H3

a(5) = &H45     'CMD

a(6) = &H2      'DATA1 : 0 :현상태 유지 1: Start 2:Stop BMP Display
a(7) = &H2      'DATA2 : 0 :현상태 유지 1: Start 2:Stop 긴급문구 Display

a(8) = &H10     'DLE     '
a(9) = &H3      'ETX

'With Jung
'If (.MSCommDisp(Port).PortOpen) Then
'    .MSCommDisp(Port).Output = a
'End If
'End With
End Sub

Public Sub GL_FrameSetting(Port As Byte)
ReDim a(9) As Byte

a(0) = &H10     'DLE
a(1) = &H2      'STX
a(2) = &H0      'DST

a(3) = &H0      'LEN
a(4) = &H2

a(5) = &H4A     'CMD

a(6) = &H2      '2단
a(7) = &H6      '6열

a(8) = &H10     'DLE
a(9) = &H3      'ETX

'With Jung
'If (.MSCommDisp(Port).PortOpen) Then
'    .MSCommDisp(Port).Output = a
'End If
'End With
End Sub

Public Sub GL_TimeSetting(Port As Byte)
ReDim a(14) As Byte

a(0) = &H10     'DLE
a(1) = &H2      'STX
a(2) = &H0      'DST

a(3) = &H0      'LEN
a(4) = &H8

a(5) = &H47     'CMD

a(6) = &H10     '10 :2010년
a(7) = &H2      '2월
a(8) = &H13     '13일
a(9) = &H6      '토요일
a(10) = &H13    '13시
a(11) = &H10    '10분
a(12) = &H30    '30초

a(13) = &H10    'DLE
a(14) = &H3     'ETX

'With Jung
'If (.MSCommDisp(Port).PortOpen) Then
'    .MSCommDisp(Port).Output = a
'End If
'End With
End Sub

Public Sub GL_Relay1(Port As Byte)
ReDim a(9) As Byte

a(0) = &H10     'DLE
a(1) = &H2      'STX
a(2) = &H0      'DST

a(3) = &H0      'LEN
a(4) = &H3

a(5) = &H4E     'CMD

a(6) = &H5     ' 1st Port 5초간 ON
a(7) = &H0     ' 2nd Port OFF

a(8) = &H10    'DLE
a(9) = &H3     'ETX

'With Jung
'If (.MSCommDisp(Port).PortOpen) Then
'    .MSCommDisp(Port).Output = a
'End If
'End With
End Sub

Public Sub GL_Relay2(Port As Byte)
ReDim a(9) As Byte

a(0) = &H10     'DLE
a(1) = &H2      'STX
a(2) = &H0      'DST

a(3) = &H0      'LEN
a(4) = &H3

a(5) = &H4E     'CMD

a(6) = &H0     ' 1st Port OFF
a(7) = &H5     ' 2nd Port 5초간 ON

a(8) = &H10    'DLE
a(9) = &H3     'ETX

'With Jung
'If (.MSCommDisp(Port).PortOpen) Then
'    .MSCommDisp(Port).Output = a
'End If
'End With
End Sub

Public Sub GL_SingnalOutputControl(Port As Byte)
ReDim a(9) As Byte

a(0) = &H10     'DLE
a(1) = &H2      'STX
a(2) = &H0      'DST

a(3) = &H0      'LEN
a(4) = &H2

a(5) = &H4E     'CMD

a(6) = &H5      ' 1st 출력신호 5초간 유지
a(7) = &HF0     ' 2nd 출력신호 항상 ON

a(8) = &H10    'DLE
a(9) = &H3     'ETX

'With Jung
'If (.MSCommDisp(Port).PortOpen) Then
'    .MSCommDisp(Port).Output = a
'End If
'End With
End Sub

Public Sub GL_Clear(Port As Byte)
ReDim GloDisp_BData(18) As Byte

GloDisp_BData(0) = &H10
GloDisp_BData(1) = &H2
GloDisp_BData(2) = &H0
GloDisp_BData(3) = &H0
GloDisp_BData(4) = &HC

GloDisp_BData(5) = &H53
GloDisp_BData(6) = &H0
GloDisp_BData(7) = &H3F
GloDisp_BData(8) = &H0
GloDisp_BData(9) = &H97

GloDisp_BData(10) = &H0
GloDisp_BData(11) = &H0
GloDisp_BData(12) = &H1
GloDisp_BData(13) = &H1
GloDisp_BData(14) = &H1E

GloDisp_BData(15) = &H2
GloDisp_BData(16) = &H1
GloDisp_BData(17) = &H10
GloDisp_BData(18) = &H3


'With Jung
'    Select Case Glo_Disp_Comm_Mode
'           Case "0" 'Tcp Ip
'                    If (.Disp_sock.State <> sckClosed) Then
'                        .Disp_sock.Close
'                    End If
'                    .Disp_sock.Connect Glo_Disp_IP, Glo_Disp_PORT
'                    Call sOutput(Glo_Disp_IP & " [DISP 접속]  시도 ")
'                    Call Err_doc("    [DISP 접속]  시도 IP = " & Glo_Disp_IP & "    PORT = " & Glo_Disp_PORT)
'           Case "1" 'UDP
'                    .Disp_sock.SendData GloDisp_BData
'                    Call sOutput(Glo_Disp_IP & " [DISP ]  UDP 전송 ")
'                    Call Err_doc("    [DISP UDP 전송]  IP = " & Glo_Disp_IP & "    PORT = " & Glo_Disp_PORT)
'           Case "2" 'Serial Rs-232c
'                    If (.MSCommDisp(0).PortOpen = True) Then
'                        .MSCommDisp(1).Output = GloDisp_BData
'                        Call sOutput(Glo_Disp_IP & " [DISP ] RS-232c 전송 ")
'                        Call Err_doc("    [DISP ] RS-232c 전송 IP = " & Glo_Disp_IP & "    PORT = " & Glo_Disp_PORT)
'                    End If
'    End Select
'End With



End Sub


Public Sub GL_Nomal(D1 As String, D2 As String, Nomal_Show As Byte, Nomal_Speed As Byte, Nomal_StopTime As Byte, Nomal_Up_color As Byte, Nomal_Down_color As Byte, IN_OUT As Integer)
    Dim Header(16) As Byte
    Dim Color_Up() As Byte
    Dim Color_Down() As Byte
    Dim Finish(1) As Byte
    Dim k1() As Byte
    Dim k2() As Byte
    Dim D() As Byte
    Dim First_Len As Integer
    Dim Second_Len As Integer
    Dim Bigger_Len As Integer
    Dim Gap_Len As Integer
    Dim i As Integer
    Dim g As Integer
    Dim First_Str As String
    Dim Second_Str As String

First_Len = LenH(D1)
Second_Len = LenH(D2)
If First_Len > Second_Len Then
    Bigger_Len = First_Len
Else
    Bigger_Len = Second_Len
End If

i = Bigger_Len Mod 12

If i = 0 Then
    
Else
    Bigger_Len = Bigger_Len + (12 - i)
End If
    
Gap_Len = Bigger_Len - First_Len
For g = 1 To Gap_Len
    D1 = D1 + " "
Next g

Gap_Len = Bigger_Len - Second_Len
For g = 1 To Gap_Len
    D2 = D2 + " "
Next g

Bigger_Len = Bigger_Len - 1

ReDim Color_Up(Bigger_Len) As Byte
ReDim Color_Down(Bigger_Len) As Byte

On Error GoTo Err_P

Header(0) = &H10   'DLE
Header(1) = &H2    'STX
Header(2) = &H0    'DST
Header(3) = &H0     'LEN
Header(4) = ((Bigger_Len + 1) * 4 + 12)
Header(5) = &H53    'CMD : 긴급 54 / 일반 53
Header(6) = &H0     'Dummy
Header(7) = &H0     'Dummy
Header(8) = &H0     '저장매체 플래시롬
Header(9) = &H91    ' (1001 0001) B[1:0] - 메인화면 폰트크기 16 font / B[5:4] - 화면표출 ON / B[6:7] - 문구 표출 방향 2 = 가로방향
Header(10) = &H0    '모듈 분할 하지 않음
Header(11) = &H0    'Dummy
Header(12) = &H0    '분할화면 효과값 : 효과없슴
Header(13) = Nomal_Show         '&H1    ' 메인화면 효과값 : 왼쪽이동
Header(14) = Nomal_Speed        '&H1E   '효과 속도
Header(15) = Nomal_StopTime     '&H0    '정지 시간 없음
Header(16) = &H0    '세로 표출 위치 : 0 행
Dim Up_Color As Byte
Dim Down_Color As Byte

Select Case Nomal_Up_color
    Case 0
        Up_Color = &H31
    Case 1
        Up_Color = &H32
    Case 2
        Up_Color = &H33
End Select
        
Select Case Nomal_Down_color
    Case 0
        Down_Color = &H31
    Case 1
        Down_Color = &H32
    Case 2
        Down_Color = &H33
End Select
        
For i = 0 To Bigger_Len
    Color_Up(i) = Up_Color    ' &H31 : 적색 / 32 : 녹색 / 33 : 노란색
Next i

For i = 0 To Bigger_Len
    Color_Down(i) = Down_Color    '&H31 : 적색 / 32 : 녹색 / 33 : 노란색
Next i
Dim D_Size1 As Integer
Dim D_Size2 As Integer

ReDim k1(Bigger_Len) As Byte
ReDim k2(Bigger_Len) As Byte

First_Str = D1
D = StrConv(First_Str, vbFromUnicode)
Bigger_Len = UBound(D)

For i = 0 To (Bigger_Len)
    k1(i) = "&H" & Hex(D(i))
Next i

Second_Str = D2
D = StrConv(Second_Str, vbFromUnicode)
Bigger_Len = UBound(D)

For i = 0 To (Bigger_Len)
    k2(i) = "&H" & Hex(D(i))
Next i
Finish(0) = &H10
Finish(1) = &H3

Dim data_len  As Integer
data_len = UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + UBound(Finish) + 5
ReDim GloDisp_BData(data_len) As Byte
For i = 0 To UBound(Header)
   GloDisp_BData(i) = Header(i)
Next i
For i = 0 To UBound(Color_Up)
    GloDisp_BData(i + UBound(Header) + 1) = Color_Up(i)
Next i
For i = 0 To UBound(Color_Down)
    GloDisp_BData(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
Next i
For i = 0 To UBound(k1)
    GloDisp_BData(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
Next i
For i = 0 To UBound(k2)
    GloDisp_BData(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
Next i
For i = 0 To UBound(Finish)
    GloDisp_BData(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
Next i

With FrmTcpServer
    Select Case IN_OUT
        Case 0
            Select Case LANE1_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock.State <> sckClosed) Then
                                .Disp1_sock.Close
                            End If
                            .Disp1_sock.Connect LANE1_DeviceIP, LANE1_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_DispPort)
                   
                   Case "1" 'UDP
                            .Disp1_sock.SendData GloDisp_BData
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_DispPort)
                   
                   Case "2" 'Serial Rs-232c
                            If (.MSCommDisp(0).PortOpen = True) Then
                                .MSCommDisp(0).Output = GloDisp_BData
                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE1_RelayComPort)
                            End If
            End Select
        
        Case 1
            Select Case LANE2_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp2_sock.State <> sckClosed) Then
                                .Disp2_sock.Close
                            End If
                            .Disp2_sock.Connect LANE2_DeviceIP, LANE2_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_DispPort)
                   
                   Case "1" 'UDP
                            .Disp2_sock.SendData GloDisp_BData
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_DispPort)
                   
                   Case "2" 'Serial Rs-232c
                            If (.MSCommDisp(1).PortOpen = True) Then
                                .MSCommDisp(1).Output = GloDisp_BData
                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE2_RelayComPort)
                            End If
            End Select
    
        Case 2
             Select Case LANE3_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp3_sock.State <> sckClosed) Then
                                .Disp3_sock.Close
                            End If
                            .Disp3_sock.Connect LANE3_DeviceIP, LANE3_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_DispPort)
                   
                   Case "1" 'UDP
                            .Disp3_sock.SendData GloDisp_BData
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_DispPort)
                   
                   Case "2" 'Serial Rs-232c
                            If (.MSCommDisp(2).PortOpen = True) Then
                                .MSCommDisp(2).Output = GloDisp_BData
                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE3_RelayComPort)
                            End If
            End Select
        
        Case 3
            Select Case LANE4_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp4_sock.State <> sckClosed) Then
                                .Disp4_sock.Close
                            End If
                            .Disp4_sock.Connect LANE4_DeviceIP, LANE4_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_DispPort)
                   
                   Case "1" 'UDP
                            .Disp4_sock.SendData GloDisp_BData
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_DispPort)
                   
                   Case "2" 'Serial Rs-232c
                            If (.MSCommDisp(3).PortOpen = True) Then
                                .MSCommDisp(3).Output = GloDisp_BData
                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE4_RelayComPort)
                            End If
            End Select
    End Select
End With

Exit Sub

Err_P:



End Sub


'긴급문구
Public Sub GL_Emergency(D1 As String, D2 As String, Led_Show As Byte, Led_Speed As Byte, Led_StopTime As Byte, Led_Repeat As Byte, Led_up_color As Byte, Led_down_color As Byte, IN_OUT As Integer)
    Dim Header(16) As Byte
    Dim Color_Up() As Byte
    Dim Color_Down() As Byte
    Dim Finish(1) As Byte
    Dim k1() As Byte
    Dim k2() As Byte
    Dim D() As Byte
    Dim First_Len As Integer
    Dim Second_Len As Integer
    Dim Bigger_Len As Integer
    Dim Gap_Len As Integer
    Dim i, g As Integer
    Dim First_Str As String
    Dim Second_Str As String
    Dim Up_Color As Byte
    Dim Down_Color As Byte
    Dim D_Size1 As Integer
    Dim D_Size2 As Integer

On Error GoTo Err_P

    First_Len = LenH(D1)
    Second_Len = LenH(D2)
    
    If First_Len > Second_Len Then
        Bigger_Len = First_Len
    Else
        Bigger_Len = Second_Len
    End If
    
    i = Bigger_Len Mod 12
    If i = 0 Then
    Else
        Bigger_Len = Bigger_Len + (12 - i)
    End If
    Gap_Len = Bigger_Len - First_Len
    For g = 1 To Gap_Len
        D1 = D1 + " "
    Next g
    Gap_Len = Bigger_Len - Second_Len
    For g = 1 To Gap_Len
        D2 = D2 + " "
    Next g
    Bigger_Len = Bigger_Len - 1
    
    ReDim Color_Up(Bigger_Len) As Byte
    ReDim Color_Down(Bigger_Len) As Byte

    Header(0) = &H10                'DLE
    Header(1) = &H2                 'STX
    Header(2) = &H0                 'DST
    Header(3) = &H0                 'LEN
    Header(4) = ((Bigger_Len + 1) * 4 + 12)
    Header(5) = &H54                'CMD : 긴급 54 / 일반 53
    Header(6) = &H0                 'Dummy
    Header(7) = &H1                 '기존 메세지 삭제 후 표출
    Header(8) = Led_Repeat          '반복 횟수
    Header(9) = &H91                ' (1001 0001) B[1:0] - 메인화면 폰트크기 16 font / B[5:4] - 화면표출 ON / B[6:7] - 문구 표출 방향 2 = 가로방향
    Header(10) = &H0                '모듈 분할 하지 않음
    Header(11) = &H0                'Dummy
    Header(12) = &H0                '분할화면 효과값 : 효과없슴
    Header(13) = Led_Show           '&H1    ' 메인화면 효과값 : 왼쪽이동
    Header(14) = Led_Speed          '&H1E   '효과 속도
    Header(15) = Led_StopTime       '&H0    '정지 시간 없음
    Header(16) = &H0                '세로 표출 위치 : 0 행


    Select Case Led_up_color
        Case 0
            Up_Color = &H31
        Case 1
            Up_Color = &H32
        Case 2
            Up_Color = &H33
    End Select
    Select Case Led_down_color
        Case 0
            Down_Color = &H31
        Case 1
            Down_Color = &H32
        Case 2
            Down_Color = &H33
    End Select
    For i = 0 To Bigger_Len
        Color_Up(i) = Up_Color    ' &H31 : 적색 / 32 : 녹색 / 33 : 노란색
    Next i
    For i = 0 To Bigger_Len
        Color_Down(i) = Down_Color    '&H31 : 적색 / 32 : 녹색 / 33 : 노란색
    Next i
    
    ReDim k1(Bigger_Len) As Byte
    ReDim k2(Bigger_Len) As Byte

    First_Str = D1
    D = StrConv(First_Str, vbFromUnicode)
    Bigger_Len = UBound(D)
    For i = 0 To (Bigger_Len)
        k1(i) = "&H" & Hex(D(i))
    Next i
    Second_Str = D2
    D = StrConv(Second_Str, vbFromUnicode)
    Bigger_Len = UBound(D)
    For i = 0 To (Bigger_Len)
        k2(i) = "&H" & Hex(D(i))
    Next i
    Finish(0) = &H10
    Finish(1) = &H3

    Dim data_len  As Integer
    data_len = UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + UBound(Finish) + 5
    ReDim GloDisp_BData(data_len) As Byte
    For i = 0 To UBound(Header)
       GloDisp_BData(i) = Header(i)
    Next i
    For i = 0 To UBound(Color_Up)
        GloDisp_BData(i + UBound(Header) + 1) = Color_Up(i)
    Next i
    For i = 0 To UBound(Color_Down)
        GloDisp_BData(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
    Next i
    For i = 0 To UBound(k1)
        GloDisp_BData(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
    Next i
    For i = 0 To UBound(k2)
        GloDisp_BData(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
    Next i
    For i = 0 To UBound(Finish)
        GloDisp_BData(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
    Next i

With FrmTcpServer
    Select Case IN_OUT
        Case 0
            Select Case LANE1_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock.State <> sckClosed) Then
                                .Disp1_sock.Close
                            End If
                            .Disp1_sock.Connect LANE1_DeviceIP, LANE1_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_DispPort)
                   
                   Case "1" 'UDP
                            .Disp1_sock.SendData GloDisp_BData
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_DispPort)
                   
                   Case "2" 'Serial Rs-232c
                            If (.MSCommDisp(0).PortOpen = True) Then
                                .MSCommDisp(0).Output = GloDisp_BData
                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE1_DispComPort)
                            End If
            End Select
        
        Case 1
            Select Case LANE2_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp2_sock.State <> sckClosed) Then
                                .Disp2_sock.Close
                            End If
                            .Disp2_sock.Connect LANE2_DeviceIP, LANE2_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_DispPort)
                   
                   Case "1" 'UDP
                            .Disp2_sock.SendData GloDisp_BData
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_DispPort)
                   
                   Case "2" 'Serial Rs-232c
                            If (.MSCommDisp(1).PortOpen = True) Then
                                .MSCommDisp(1).Output = GloDisp_BData
                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE2_DispComPort)
                            End If
            End Select
    
        Case 2
            Select Case LANE3_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp3_sock.State <> sckClosed) Then
                                .Disp3_sock.Close
                            End If
                            .Disp3_sock.Connect LANE3_DeviceIP, LANE3_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_DispPort)
                   
                   Case "1" 'UDP
                            .Disp3_sock.SendData GloDisp_BData
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_DispPort)
                   
                   Case "2" 'Serial Rs-232c
                            If (.MSCommDisp(2).PortOpen = True) Then
                                .MSCommDisp(2).Output = GloDisp_BData
                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE3_DispComPort)
                            End If
            End Select
        
        
        Case 3
            Select Case LANE4_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp4_sock.State <> sckClosed) Then
                                .Disp4_sock.Close
                            End If
                            .Disp4_sock.Connect LANE4_DeviceIP, LANE4_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_DispPort)
                   
                   Case "1" 'UDP
                            .Disp4_sock.SendData GloDisp_BData
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_DispPort)
                   
                   Case "2" 'Serial Rs-232c
                            If (.MSCommDisp(3).PortOpen = True) Then
                                .MSCommDisp(3).Output = GloDisp_BData
                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE4_DispComPort)
                            End If
            End Select
    
    End Select

End With

Exit Sub
Err_P:


End Sub

'상황판 일반문구

Public Sub GL_BIG(Message As String, Color As String, SEffect As Byte, SSpeed As Byte, STime As Byte, SRepeat As Byte, IN_OUT As Byte)

'Dim Header(16) As Byte
'Dim Color_A(63) As Byte
'Dim str(63) As Byte
'
'Dim Finish(1) As Byte
'
'Dim i As Integer
'Dim s As Integer
'Dim q As Integer
'
'Dim D() As Byte
'
'With Jung
'
'Header(0) = &H10    'DLE
'Header(1) = &H2     'STX
'Header(2) = &H0     'DST
'Header(3) = &H0     'LEN
'Header(4) = &H8C    '140 Byte
'Header(5) = &H53    'CMD : 긴급 54 / 일반 53
'Header(6) = &H0     'Dummy
'Header(7) = &H0     'Dummy
'Header(8) = &H0     '53 : 저장매체 00 플래시 01 램 54:SRepeat    '반복횟수
'Header(9) = &H57    ' (0101 0111) B[1:0] - 메인화면 폰트크기 16 font / B[5:4] - 화면표출 ON / B[6:7] - 문구 표출 방향 2 = 가로방향
'Header(10) = &H0    '모듈 분할 하지 않음
'Header(11) = &H0    'Dummy
'Header(12) = &H0    '분할화면 효과값 : 효과없슴
'Header(13) = &H0    'ShowEffect         '&H1    ' 메인화면 효과값 : 왼쪽이동
'Header(14) = &H14   'ShowSpeed        '&H1E   '효과 속도
'Header(15) = &HFA   'STime     '&H0    '정지 시간 없음
'Header(16) = &H0    '세로 표출 위치 : 0 행
'
'
'For i = 1 To 8
'    C = Mid(Color, i, 1)
'    If C = "G" Then
'        For s = 1 To 8
'            q = s + ((i - 1) * 8)
'            Color_A(q - 1) = &H32
'        Next s
'    ElseIf C = "R" Then
'        For s = 1 To 8
'            q = s + ((i - 1) * 8)
'            Color_A(q - 1) = &H31
'        Next s
'    ElseIf C = "O" Then
'        For s = 1 To 8
'            q = s + ((i - 1) * 8)
'            Color_A(q - 1) = &H33
'        Next s
'    End If
'Next i
'
'D = StrConv(Message, vbFromUnicode)
'
'For i = 0 To 63
'    str(i) = "&H" & Hex(D(i))
'Next i
'
'
'Finish(0) = &H10
'Finish(1) = &H3
'
'If (.MSCommDisp(IN_OUT).PortOpen) Then
'    .MSCommDisp(IN_OUT).Output = Header
'    .MSCommDisp(IN_OUT).Output = Color_A
'    .MSCommDisp(IN_OUT).Output = str
'    .MSCommDisp(IN_OUT).Output = Finish
'End If
'
''Exit Sub
'
''Dim start As Single
''start = Timer
''Do While Timer < start + 0.2
''    DoEvents
''    If (Timer < start) Then
''        start = start - 86400
''    End If
''    If (In_Led_F = True) Then
''        Exit Sub
''    End If
''Loop
''Call LED_Int_On(IN_OUT)
'
'End With
'
'Exit Sub
'err_P:


End Sub

'상황판 긴급문구 테스트용

Public Sub GL_BIG_Test(Message As String, IN_OUT As Byte)

'Dim E(146) As Byte
'Dim i As Integer
'Dim s As Integer
'Dim D(146) As Integer
'
'With Jung
'
'For i = 0 To 146
'
'    s = (i * 3) + 1
'    E(i) = "&H" + Mid(Message, s, 2)
'Next i
'
'
'If (.MSCommDisp(IN_OUT).PortOpen) Then
'    .MSCommDisp(IN_OUT).Output = E
'End If
'
'End With
'
'Exit Sub
'err_P:


End Sub

Public Sub DFee(FeeVal As String)
'Dim tmp As String
'Dim a() As Byte
'ReDim a(12)
'Dim i As Integer
'Dim DLen As Integer
'Dim ChkSum As Integer
'
'With Jung
'    tmp = CStr(FeeVal)
'    If (Len(tmp) > 6) Then
'        tmp = "999999"
'    Else
'        tmp = Space(6 - Len(tmp)) & tmp
'    End If
'
'    If (.MSCommDisp(0).PortOpen) Then
'    Else
'        Exit Sub
'    End If
'    .MSCommDisp(0).Output = Chr$(2) & "WD" & tmp & Chr$(3)
'End With
End Sub


Public Sub Relay_Out(RNum As Integer, GateNo As Integer)

'RNum 0 : Gate Relay, 1: Capture Test


With FrmTcpServer
    If (RNum = 0) Then
        GlO_TcpDataGate = Chr$(2) & "R2" & Chr$(3)
    Else
        GlO_TcpDataGate = Chr$(2) & "R1" & Chr$(3)
    End If
    
    Select Case GateNo
        Case 0
            Select Case LANE1_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Gate1_sock.State <> sckClosed) Then
                                .Gate1_sock.Close
                            End If
                            .Gate1_sock.Connect LANE1_DeviceIP, LANE1_RelayPort
                            If (RNum = 0) Then
                                Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                            Else
                                Call DataLogger("[Get Frame TCP/IP 접속]  시도 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                            End If
                            Call None_Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Gate1_sock.SendData (GlO_TcpDataGate)
                            If (RNum = 0) Then
                                Call DataLogger("[GATE UDP 전송]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                            Else
                                Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                            End If
                            Call None_Delay_Time(0.1)
                   
                   Case "2" 'Serial Rs-232c
                            If (.MSCommGate(0).PortOpen = True) Then
                                .MSCommGate(0).Output = GlO_TcpDataGate
                                If (RNum = 0) Then
                                    Call DataLogger("[GATE ] RS-232c 전송 COM PORT = " & LANE1_RelayComPort)
                                Else
                                    Call DataLogger("[Get Frame] RS-232c 전송 COM PORT = " & LANE1_RelayComPort)
                                End If
                            End If
            End Select
        
        Case 1
            Select Case LANE2_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Gate2_sock.State <> sckClosed) Then
                                .Gate2_sock.Close
                            End If
                            .Gate2_sock.Connect LANE2_DeviceIP, LANE2_RelayPort
                            If (RNum = 0) Then
                                Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                            Else
                                Call DataLogger("[Get Frame TCP/IP 접속]  시도 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                            End If
                            Call None_Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Gate2_sock.SendData (GlO_TcpDataGate)
                            If (RNum = 0) Then
                                Call DataLogger("[GATE UDP 전송]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                            Else
                                Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                            End If
                            Call None_Delay_Time(0.1)
                   
                   Case "2" 'Serial Rs-232c
                            If (.MSCommGate(1).PortOpen = True) Then
                                .MSCommGate(1).Output = GlO_TcpDataGate
                                If (RNum = 0) Then
                                    Call DataLogger("[GATE ] RS-232c 전송 COM PORT = " & LANE2_RelayComPort)
                                Else
                                    Call DataLogger("[Get Frame] RS-232c 전송 COM PORT = " & LANE2_RelayComPort)
                                End If
                            End If
            End Select
        
        Case 2
            Select Case LANE3_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Gate3_sock.State <> sckClosed) Then
                                .Gate3_sock.Close
                            End If
                            .Gate3_sock.Connect LANE3_DeviceIP, LANE3_RelayPort
                            If (RNum = 0) Then
                                Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                            Else
                                Call DataLogger("[Get Frame TCP/IP 접속]  시도 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                            End If
                            Call None_Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Gate3_sock.SendData (GlO_TcpDataGate)
                            If (RNum = 0) Then
                                Call DataLogger("[GATE UDP 전송]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                            Else
                                Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                            End If
                            Call None_Delay_Time(0.1)
                   
                   Case "2" 'Serial Rs-232c
                            If (.MSCommGate(2).PortOpen = True) Then
                                .MSCommGate(2).Output = GlO_TcpDataGate
                                If (RNum = 0) Then
                                    Call DataLogger("[GATE ] RS-232c 전송 COM PORT = " & LANE3_RelayComPort)
                                Else
                                    Call DataLogger("[Get Frame] RS-232c 전송 COM PORT = " & LANE3_RelayComPort)
                                End If
                            End If
            End Select
        
        Case 3
            Select Case LANE4_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Gate4_sock.State <> sckClosed) Then
                                .Gate4_sock.Close
                            End If
                            .Gate4_sock.Connect LANE4_DeviceIP, LANE4_RelayPort
                            If (RNum = 0) Then
                                Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                            Else
                                Call DataLogger("[Get Frame TCP/IP 접속]  시도 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                            End If
                            Call None_Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Gate4_sock.SendData (GlO_TcpDataGate)
                            If (RNum = 0) Then
                                Call DataLogger("[GATE UDP 전송]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                            Else
                                Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                            End If
                            Call None_Delay_Time(0.1)
                   
                   Case "2" 'Serial Rs-232c
                            If (.MSCommGate(3).PortOpen = True) Then
                                .MSCommGate(3).Output = GlO_TcpDataGate
                                If (RNum = 0) Then
                                    Call DataLogger("[GATE ] RS-232c 전송 COM PORT = " & LANE4_RelayComPort)
                                Else
                                    Call DataLogger("[Get Frame] RS-232c 전송 COM PORT = " & LANE4_RelayComPort)
                                End If
                            End If
            End Select
    
    End Select
    
    
End With

End Sub


