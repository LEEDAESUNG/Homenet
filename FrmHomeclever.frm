VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmHomeclever 
   BorderStyle     =   1  '단일 고정
   Caption         =   "홈클래버 홈넷 프로그램"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   17430
   StartUpPosition =   3  'Windows 기본값
   Begin MSWinsockLib.Winsock HomeR_Sock 
      Left            =   2655
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   " 세대통보 테스트 "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   11820
      TabIndex        =   10
      Top             =   600
      Width           =   5325
      Begin VB.TextBox txt_Dong 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Text            =   "0101"
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txt_Ho 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Text            =   "0202"
         Top             =   330
         Width           =   915
      End
      Begin VB.CommandButton cmd_Test 
         Caption         =   "세대통보 테스트"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3210
         TabIndex        =   11
         Top             =   300
         Width           =   1875
      End
      Begin VB.Label Label2 
         Caption         =   "동"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1260
         TabIndex        =   15
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "호"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   2730
         TabIndex        =   14
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   6360
   End
   Begin VB.CommandButton cmd_Clear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16110
      TabIndex        =   8
      Top             =   6450
      Width           =   1065
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4560
      Left            =   120
      TabIndex        =   0
      Top             =   1710
      Width           =   17220
   End
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "종 료"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   16050
      TabIndex        =   7
      Top             =   150
      Width           =   1095
   End
   Begin VB.TextBox txt_HostPort 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      TabIndex        =   6
      Text            =   "12121"
      Top             =   1050
      Width           =   1755
   End
   Begin VB.TextBox txt_HomeNet_Port 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      TabIndex        =   4
      Text            =   "12121"
      Top             =   540
      Width           =   1755
   End
   Begin VB.TextBox txt_HomeNet_IP 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      TabIndex        =   3
      Text            =   "127.0.0.1"
      Top             =   180
      Width           =   1755
   End
   Begin MSWinsockLib.Winsock HomeSock 
      Left            =   630
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock HostSock 
      Left            =   180
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label lbl_Date 
      Caption         =   "Timer"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9870
      TabIndex        =   9
      Top             =   240
      Width           =   3075
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   17310
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Host RCV Port :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   270
      TabIndex        =   5
      Top             =   1080
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "HomeNet Port :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   270
      TabIndex        =   2
      Top             =   570
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "HomeNet IP :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   1
      Top             =   210
      Width           =   1305
   End
End
Attribute VB_Name = "FrmHomeclever"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Reconn_Count As Integer

Private Sub cmd_Clear_Click()
    List1.Clear
End Sub

Private Sub cmd_Exit_Click()
    Call Form_QueryUnload(0, 0)
End Sub

Private Sub cmd_Test_Click()
    If Len(txt_Dong.Text) <> 0 And Len(txt_Ho.Text) <> 0 Then
        Call Homeclever_Proc(Trim(txt_Dong.Text), Trim(txt_Ho.Text), "01가1234")
    Else
        MsgBox ("테스트할 동/호를 확인하세요..!!")
    End If
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then
        End
    End If
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    
    IniFileName$ = App.Path & "\HomeNet.ini"
    Doc_Path_Name$ = App.Path & "\Doc\"
    
    txt_HostPort.Text = HostPort
    txt_HomeNet_IP.Text = HomeNet_IP
    txt_HomeNet_Port.Text = HomeNet_Port
    
    '호스트 데이터 수신
    HostSock.Protocol = sckUDPProtocol
    HostSock.LocalPort = HostPort
    HostSock.Bind
    
    '홈넷 데이터 수신
    HomeR_Sock.Protocol = sckTCPProtocol
    HomeR_Sock.LocalPort = 55502
    HomeR_Sock.Listen
    
    Call HomeLogger("[HomeNet Program ] Homeclever HomeNet Start..!!")
    
    Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msg, Style, Title, Response
    Dim Ret As Boolean
    msg = " 홈넷 프로그램을 종료하시겠습니까?         "
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = " Parking Manager™ "
    Response = MsgBox(msg, Style, Title)
    If Response = vbYes Then
        
        Call HomeLogger("[HomeNet Exit Proc]    " & " HomeNet Program END..!!")
        End
    End If
    Me.MousePointer = 0
    Cancel = True
End Sub

Private Sub Timer1_Timer()
    lbl_Date.Caption = Format(Now, "YYYY-MM-DD HH:NN:SS")
    
    If Format(Now, "NNSS") = "0001" Then
        List1.Clear
    End If
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HOST로부터 Home_UDP 받기
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HostSock_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Dim Tmp_Path As String
    Dim i, GateNo As Integer
    Dim CarNum As String
    
On Error GoTo Err_P
    
    HostSock.GetData sData, , bytesTotal
    Call HomeLogger("HostSock UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sData)
    
    Call HomeNet_Proc(sData)
    
    Call Homeclever_Proc(HomeNet_Dong, HomeNet_Ho, HomeNet_CarNo)

Exit Sub

Err_P:
    Call HomeLogger(" [HostSock UDP DataArrival]  " & Err.Description)

End Sub
Private Sub HostSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Call HomeLogger(" [HostSock Error]  " & Description)
End Sub



Public Sub Homeclever_Proc(Dong As String, Ho As String, CarNo As String)
    ReDim HC_Tcp_Header(47) As Byte
    Dim tmpHeaderDong() As Byte
    Dim tmpHeaderHo() As Byte
    Dim tmpHeaderCarno() As Byte
    Dim Socket_Data_len As String
    Dim tmpDong As String
    Dim tmpHo As String
    Dim tmpCarNo As String
    
On Error GoTo Err_Proc
    
    '예) 106동 1203호 서울3다3246 ==> 0xF2003001106동     1203호    서울3다3246         0xF3
    '데이터전송프로토콜(48byte)
    '1byte:0xF2
    '2byte:기기번호
    '4byte:Command(3001:데이터전송, 4101:데이터수신OK, 4102:데이터수신완료 Err 및 데이터 재전송요구)
    '10byte:동
    '10byte:호수
    '20byte;차량번호
    '1byte:0xF3
    tmpDong = Format(Dong, "0") & "동": tmpDong = tmpDong & Space(10 - LenH(tmpDong))
    tmpHo = Format(Ho, "0") & "호":     tmpHo = tmpHo & Space(10 - LenH(tmpHo))
    tmpCarNo = CarNo:                   tmpCarNo = tmpCarNo & Space(20 - LenH(tmpCarNo))

    HC_Tcp_Header(0) = &HF2
    HC_Tcp_Header(1) = Asc("0")
    HC_Tcp_Header(2) = Asc("1")
    HC_Tcp_Header(3) = Asc("3")
    HC_Tcp_Header(4) = Asc("0")
    HC_Tcp_Header(5) = Asc("0")
    HC_Tcp_Header(6) = Asc("1")

    tmpHeaderDong = StrConv(tmpDong, vbFromUnicode)
    i = 0
    For i = 0 To 10 - 1
        HC_Tcp_Header(7 + i) = tmpHeaderDong(i)
    Next i

    tmpHeaderHo = StrConv(tmpHo, vbFromUnicode)
    For i = 0 To 10 - 1
        HC_Tcp_Header(17 + i) = tmpHeaderHo(i)
    Next i

    tmpHeaderCarno = StrConv(tmpCarNo, vbFromUnicode)
    For i = 0 To 20 - 1
        HC_Tcp_Header(27 + i) = tmpHeaderCarno(i)
    Next i

    HC_Tcp_Header(47) = &HF3
    
    
    Reconn_Count = 0 '재접속 카운트
    Call Socket_Connect

Exit Sub

Err_Proc:
    Call HomeLogger("[Homeclever_Proc]  " & Err.Description)
End Sub

Public Sub Socket_Connect()
Dim bdata() As Byte

On Error GoTo Err_P
    Call HomeLogger("[Homeclever Connect] 홈넷접속시도 : " & HomeNet_IP & " " & HomeNet_Port)
    If (HomeSock.State <> sckClosed) Then
        HomeSock.Close
    End If
    
    HomeSock.Connect HomeNet_IP, HomeNet_Port

Exit Sub

Err_P:
    Call HomeLogger("[Homeclever Connect] 에러내용 : " & Err.Description)

End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 세대통보 데이터 전송처리
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HomeSock_Connect()
On Error GoTo Err_P
    HomeSock.SendData HC_Tcp_Header
Exit Sub
Err_P:
    Call HomeLogger(Format(Now, "yyyy-mm-dd hh:nn:ss") & " [HomeSock_Connect 프로시져] 에러내용 : " & Err.Description)
End Sub
Private Sub HomeSock_SendComplete()
    Dim strData As String
    strData = StrConv(HC_Tcp_Header, vbUnicode)
    Call HomeLogger("[Homeclever Send Complete] " & strData)
    'HomeSock.Close 'Close 할 경우, 홈넷서버에서 40006 에러 발생되어 주석처리함
End Sub
Private Sub HomeSock_DataArrival(ByVal bytesTotal As Long)
    Dim rMsg As String
    Dim B() As String
    Dim i As Long
    
On Error GoTo Err_P
    '실제로는 데이터 안 들어오며 아래 코드 실행안됨
    ReDim B(bytesTotal - 1)
    HomeSock.GetData B(), vbArray + vbByte, bytesTotal
    rMsg = StrConv(B, vbFromUnicode)

    Call HomeLogger("[Homeclever RCV] : " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & rMsg)
    

    For i = 0 To bytesTotal - 1
        If (B(i) >= &H80) Then
            rMsg = rMsg & Chr$(Val("&H" & Hex(B(i)) & Hex(B(i + 1))))
            i = i + 1
        Else
            rMsg = rMsg & Chr$(B(i))
        End If
    Next i
    
    If (InStr(rMsg, "4101") > 0) Then
        Call HomeLogger("[Homeclever RCV]  RCV : " & rMsg & "(데이터수신OK)")
    ElseIf (InStr(rMsg, "4102") > 0) Then
        Call HomeLogger("[Homeclever RCV]  RCV : " & rMsg & "(데이터수신완료 Err)")
    Else
        Call HomeLogger("[Homeclever RCV]  RCV : " & rMsg & "(수신데이터 Err)")
    End If
    
    Exit Sub
Err_P:
    Call DataLogger(" [Homeclever RCV] Err : " & Err.Description)
End Sub
Private Sub HomeSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call HomeLogger(" [Homeclever Error]  " & Description)
    
    HomeSock.Close
    
    Reconn_Count = Reconn_Count + 1
    If (Reconn_Count < 3) Then
        HomeSock.Connect HomeNet_IP, HomeNet_Port
    End If
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 홈넷 서버로부터 데이터 수신처리
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HomeR_Sock_ConnectionRequest(ByVal requestID As Long)
    HomeR_Sock.Close
    HomeR_Sock.Accept requestID
    Call HomeLogger("[Homeclever Svr Accept] ")
End Sub
Private Sub HomeR_Sock_DataArrival(ByVal bytesTotal As Long)
    Dim rMsg As String
    Dim B() As String
    Dim i As Long
    
On Error GoTo Err_P

    ReDim B(bytesTotal - 1)
    HomeR_Sock.GetData B(), vbArray + vbByte, bytesTotal
    rMsg = StrConv(B, vbFromUnicode)

    Call HomeLogger("[Homeclever Svr RCV] : " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & rMsg)
    

    For i = 0 To bytesTotal - 1
        If (B(i) >= &H80) Then
            rMsg = rMsg & Chr$(Val("&H" & Hex(B(i)) & Hex(B(i + 1))))
            i = i + 1
        Else
            rMsg = rMsg & Chr$(B(i))
        End If
    Next i
    
    If (InStr(rMsg, "4101") > 0) Then
        Call HomeLogger("[Homeclever_Proc]  RCV : " & rMsg & "(데이터수신OK)")
    ElseIf (InStr(rMsg, "4102") > 0) Then
        Call HomeLogger("[Homeclever_Proc]  RCV : " & rMsg & "(데이터수신완료 Err)")
    Else
        Call HomeLogger("[Homeclever_Proc]  RCV : " & rMsg & "(수신데이터 Err)")
    End If
    
    Exit Sub
Err_P:
    Call DataLogger(" [Homeclever Svr RCV] Err : " & Err.Description)
End Sub
Private Sub HomeR_Sock_Close()
    HomeR_Sock.Close
    HomeR_Sock.Listen
    
    Call HomeLogger("[Homeclever Svr Close] ")
End Sub
Private Sub HomeR_Sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MobileR_Sock.Close
    MobileR_Sock.Listen
    
    Call DataLogger("[Homeclever Svr Error] " & Description)
End Sub


