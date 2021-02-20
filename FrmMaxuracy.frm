VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMaxuracy 
   BorderStyle     =   1  '단일 고정
   Caption         =   "맥서러씨 홈넷 프로그램"
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
      ItemData        =   "FrmMaxuracy.frx":0000
      Left            =   120
      List            =   "FrmMaxuracy.frx":0002
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
Attribute VB_Name = "FrmMaxuracy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Clear_Click()
    List1.Clear
End Sub

Private Sub cmd_Exit_Click()
    Call Form_QueryUnload(0, 0)
End Sub

Private Sub cmd_Test_Click()
    If Len(txt_Dong.Text) <> 0 And Len(txt_Ho.Text) <> 0 Then
        Call Maxuracy_Proc(Trim(txt_Dong.Text), Trim(txt_Ho.Text), "01가1234")
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

HostSock.Protocol = sckUDPProtocol
HostSock.LocalPort = HostPort
HostSock.Bind

Call HomeLogger("[HomeNet Program ] Maxuracy HomeNet Start..!!")


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
    Call HomeLogger("[HOME Exit Proc]    " & " HomeNet Program END..!!")
    'Call HomeDB_Close(adoHome)
    End
End If
Me.MousePointer = 0
Cancel = True
End Sub


Private Sub HomeSock_SendComplete()
    Call HomeLogger("Send complete:" & Socket_Data)
End Sub

'HOST로부터 Home_UDP 받기
Private Sub HostSock_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Dim Tmp_Path As String
    Dim i, GateNo As Integer
    Dim CarNum As String
    
On Error GoTo Err_P
    
    HostSock.GetData sData, , bytesTotal
    Call HomeLogger("HostSock UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sData)
    
    Call HomeNet_Proc(sData)
    
    '현대통신 세대통보
    Call Maxuracy_Proc(HomeNet_Dong, HomeNet_Ho, HomeNet_CarNo)

Exit Sub

Err_P:
    Call HomeLogger(" [HostSock UDP DataArrival]  " & Err.Description)

End Sub

Private Sub HomeSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call HomeLogger(" [HomeSock Error]  " & Description)
End Sub

Private Sub HostSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Call HomeLogger(" [HostSock Error]  " & Description)
End Sub

Private Sub Timer1_Timer()
lbl_Date.Caption = Format(Now, "YYYY-MM-DD HH:NN:SS")

If Format(Now, "NNSS") = "0001" Then
    List1.Clear
End If

End Sub

Public Sub Maxuracy_Proc(Dong As String, Ho As String, CarNo As String)
    Dim Socket_Data_len As String * 8

On Error GoTo Err_Proc
    
    ' 첫바이트 0:입차, 1:출차, 응답없음.
    ' 예) 106동 1203호 서울3다3246 2018/3/29 13:24:01 입차했을 경우 ==> 0,0106,1203,서울3다3246,20180329132401
    Dong = Format(Val(Dong), "0000")
    Ho = Format(Val(Ho), "0000")
    Socket_Data = "0," & Dong & "," & Ho & "," & Trim(CarNo) & "," & Format(Now, "yyyymmddhhnnss")
    Call HomeLogger("[Maxuracy_Proc]  SND : " & Socket_Data)
    Call Socket_Connect

Exit Sub

Err_Proc:
    Call HomeLogger("[Maxuracy_Proc]  " & Err.Description)
End Sub

Public Sub Socket_Connect()
Dim bdata() As Byte

On Error GoTo Err_P
    Call HomeLogger("[HomeAlarm_SocketConnect] 홈넷접속시도 : " & HomeNet_IP & " " & HomeNet_Port)
    If (HomeSock.State <> sckClosed) Then
        HomeSock.Close
    End If
    
    HomeSock.Connect HomeNet_IP, HomeNet_Port

Exit Sub

Err_P:
    Call HomeLogger("[HomeAlarm_SocketConnect] 에러내용 : " & Err.Description)

End Sub

Private Sub HomeSock_Connect()
Dim sData As String
Dim bdata() As Byte
Dim i As Integer

On Error GoTo Err_P

    sData = Socket_Data

    'ReDim bData(Len(sdata) - 1) As Byte
    'bData = StrConv(sdata, vbNarrow)
    'bData = StrConv(sdata, vbUnicode) '한글깨짐
    'HomeSock.SendData bData
    HomeSock.SendData sData

Exit Sub

Err_P:
Call HomeLogger(Format(Now, "yyyy-mm-dd hh:nn:ss") & " [HomeSock_Connect 프로시져] 에러내용 : " & Err.Description)
End Sub

'맥서러시는 서버응답없음
Private Sub HomeSock_DataArrival(ByVal bytesTotal As Long)
Dim rMsg As String
Dim B() As Byte
Dim Ret As Integer
Dim i As Integer
Dim sData As String

On Error GoTo Err_P

ReDim B(bytesTotal - 1)

HomeSock.GetData B(), vbArray + vbByte, bytesTotal
For i = 0 To bytesTotal - 1
    If (B(i) >= &H80) Then
        rMsg = rMsg & Chr$(Val("&H" & Hex(B(i)) & Hex(B(i + 1))))
        i = i + 1
    Else
        rMsg = rMsg & Chr$(B(i))
    End If
Next i

'Call HomeLogger(" [홈네트워크 패킷 수신완료]" & " " & rMsg)
Call HomeLogger("[Maxuracy_Proc]  RCV : " & rMsg)

HomeSock.Close

Exit Sub

Err_P:
    Call HomeLogger(" [Maxuracy_Proc RCV] 에러내용 : " & Err.Description)
End Sub



