VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmKocomRS232 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "코콤(시리얼) 세대통보"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
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
      Left            =   8850
      TabIndex        =   9
      Top             =   630
      Width           =   5325
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
         TabIndex        =   12
         Top             =   300
         Width           =   1875
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
         TabIndex        =   11
         Text            =   "0202"
         Top             =   330
         Width           =   915
      End
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
         TabIndex        =   10
         Text            =   "0101"
         Top             =   330
         Width           =   915
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
         Index           =   2
         Left            =   1260
         TabIndex        =   13
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   9120
      TabIndex        =   8
      Top             =   6450
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1020
      Top             =   6270
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      ParityReplace   =   48
      RThreshold      =   1
      InputMode       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   540
      Top             =   6270
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
      Left            =   13050
      TabIndex        =   3
      Top             =   6360
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
      Height          =   4335
      ItemData        =   "frmKocomRS232.frx":0000
      Left            =   30
      List            =   "frmKocomRS232.frx":0002
      TabIndex        =   2
      Top             =   1680
      Width           =   14160
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
      Left            =   13080
      TabIndex        =   1
      Top             =   90
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
      Left            =   1710
      TabIndex        =   0
      Text            =   "12121"
      Top             =   120
      Width           =   1755
   End
   Begin MSWinsockLib.Winsock HostSock 
      Left            =   90
      Top             =   6270
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
      Left            =   9780
      TabIndex        =   7
      Top             =   180
      Width           =   3075
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   30
      X2              =   14190
      Y1              =   1650
      Y2              =   1650
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
      Left            =   180
      TabIndex        =   6
      Top             =   150
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "Kocom CNT : "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4110
      TabIndex        =   5
      Top             =   150
      Width           =   1275
   End
   Begin VB.Label lbl_KocomCnt 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5460
      TabIndex        =   4
      Top             =   150
      Width           =   1275
   End
End
Attribute VB_Name = "frmKocomRS232"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call KocomRS232_Alarm(0, "서울01가9012", "1234", "5678")
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

HostSock.Protocol = sckUDPProtocol
HostSock.LocalPort = HostPort
HostSock.Bind

Call HomeLogger("[HomeNet Program] KOCOM RS232 HomeNet Start..!!")

'세대통보 포트 열기 (서울통신)
MSComm1.CommPort = Val(Get_Ini("Kocom232", "HomeNet_ComPort", "1"))
MSComm1.Settings = "9600,n,8,1"
MSComm1.InputLen = 1
MSComm1.InputMode = comInputModeBinary
MSComm1.PortOpen = True

If (MSComm1.PortOpen = True) Then
    Call HomeLogger(" [HomeCOM Port Open Proc] Port Open Success..!! PortNo : " & MSComm1.CommPort)
Else
    Call HomeLogger(" [HomeCOM Port Open Proc] Port Open Failure..!! PortNo : " & MSComm1.CommPort)
    MsgBox "    세대통보 시리얼포트 오픈 실패..!!"
    End
End If

Timer1.Enabled = True

End Sub

Private Sub MSComm1_OnComm()
    Dim i, Cnt As Integer
    Dim tmpinstring() As Byte
    Dim Instring As String
    Dim Data(3) As Byte
    Dim bchk As Byte
    Dim RCV_KOCOM(20) As Byte
    
'On Error GoTo Err_Proc

Select Case MSComm1.CommEvent
        Case comEvDSR
        Case comEvSend
        Case comEvCTS
        Case comEvReceive
                '한바이트를 읽어온다
                If i > 20 Then
                    i = 20
                End If
                                
                tmpinstring = MSComm1.Input
                Instring = Format$(Hex(tmpinstring(i)), "00")
                
                If Instring = "AA" Then
                    Instring = "AA"
                    i = 0
                    RCV_KOCOM(0) = "&H" & Instring
                Else
                    i = i + 1
                    RCV_KOCOM(i) = "&H" & Instring
                End If
                
                If i = 20 Then
                    Select Case RCV_KOCOM(15)
                         Case 2
                                 Select Case RCV_KOCOM(16)
                                     Case 2
                                        Call HomeLogger("RCV_KOCOM(16) Error : 목적지 없음")
                                     Case 4
                                        Call HomeLogger("RCV_KOCOM(16) Error : 통신버퍼 초과")
                                     Case 6
                                        Call HomeLogger("RCV_KOCOM(16) Error : 장치기 미등록")
                                     Case Else
                                        Call HomeLogger("RCV_KOCOM(16) Error : 알수없는 데이터 : " & RCV_KOCOM(16))
                                 End Select
                         Case 9
                                 Select Case RCV_KOCOM(16)
                                     Case 1
                                        Call HomeLogger("RCV_KOCOM : 입차 차량번호 : " & CByte(RCV_KOCOM(17)) & CByte(RCV_KOCOM(18)))
                                     Case Else
                                         Call HomeLogger("RCV_KOCOM(16) : 알수없는 데이터 : " & RCV_KOCOM(16))
                                 End Select
                         Case Else
                             Call HomeLogger("RCV_KOCOM(15) : 알수없는 데이터 : " & RCV_KOCOM(15))
                    End Select
                    MSComm1.InBufferCount = 0
                End If
End Select

Exit Sub

Err_Proc:
    Call HomeLogger(" [HomeAlarm Proc] Err_Proc : " & Err.Description)

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
    
    '코콤 세대통보
    'Kocom_Alarm(inout As Integer, CarNo As String, tmpDong As Integer, tmpHo As Integer)
    Call KocomRS232_Alarm(0, Trim(HomeNet_CarNo), HomeNet_Dong, HomeNet_Ho)

Exit Sub

Err_P:
    Call HomeLogger("[HostSock UDP DataArrival]  " & Err.Description)

End Sub

Private Sub HostSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call HomeLogger("[HostSock UDP Error]  " & Description)
End Sub

Private Sub Timer1_Timer()
lbl_Date.Caption = Format(Now, "YYYY-MM-DD HH:NN:SS")

If Format(Now, "NNSS") = "0001" Then
    List1.Clear
End If

'If (Kocom_Cnt < 50) Then
'    Kocom_Cnt = Kocom_Cnt + 1
'Else
'    If Kocom_Mode = "" Or Kocom_Mode = "BIND" Then
'        Kocom_Cnt = 0
'        Call Kocom_ALIVE
'    End If
'End If
'lbl_KocomCnt.Caption = Kocom_Cnt

End Sub

Public Sub KocomRS232_Alarm(inout As Integer, CarNo As String, tmpDong As String, tmpHo As String)

    Dim i As Integer
    Dim DTime As String
    Dim SND_KOCOM(30) As Byte
    Dim CheckSum As Long
    Dim tmpCarNo As String
    
    SND_KOCOM(0) = &HAA
    SND_KOCOM(1) = &H55
    
    SND_KOCOM(2) = &HD8
    
    SND_KOCOM(3) = &HE4 '2nd &HE5, 3rd &HE6
    
    SND_KOCOM(4) = &H0  'PCNT 고정
    
    SND_KOCOM(5) = &H1   '01 주차관제에서 주장치로 송신 , 02 반대의 경우
    
    SND_KOCOM(6) = &H1
    SND_KOCOM(7) = &H1
    SND_KOCOM(8) = &H0
    SND_KOCOM(9) = &H0
    
    SND_KOCOM(10) = &H82
    
    '동 호 입력
    SND_KOCOM(11) = "&H" & Mid$(tmpDong, 1, 2)
    SND_KOCOM(12) = "&H" & Mid$(tmpDong, 3, 2)
    SND_KOCOM(13) = "&H" & Mid$(tmpHo, 1, 2)
    SND_KOCOM(14) = "&H" & Mid$(tmpHo, 3, 2)
    
    tmpCarNo = Right(CarNo, 4)
    
    SND_KOCOM(15) = &H9     'Command1
    SND_KOCOM(16) = &H1     'Command2
    SND_KOCOM(17) = "&H" & Mid$(tmpCarNo, 1, 2)
    SND_KOCOM(18) = "&H" & Mid$(tmpCarNo, 3, 2)
    
    SND_KOCOM(19) = &H1     '02 출차알림
    
    DTime = Format(Now, "YYYYMMDDHHNN")
    
    SND_KOCOM(20) = "&H" & Mid$(DTime, 1, 2)    '년 상위
    SND_KOCOM(21) = "&H" & Mid$(DTime, 3, 2)    '년 하위
    SND_KOCOM(22) = "&H" & Mid$(DTime, 5, 2)     '월
    SND_KOCOM(23) = "&H" & Mid$(DTime, 7, 2)     '일
    SND_KOCOM(24) = "&H" & Mid$(DTime, 9, 2)    '시
    SND_KOCOM(25) = "&H" & Mid$(DTime, 11, 2)    '분
    SND_KOCOM(26) = &H0     '예비
    SND_KOCOM(27) = &H0     '예비
    SND_KOCOM(28) = &H0     '예비
    
    CheckSum = 0
    For i = 2 To 28
        CheckSum = CheckSum + SND_KOCOM(i)
        If CheckSum > 255 Then CheckSum = CheckSum - 256
    Next i
    SND_KOCOM(29) = CByte(CheckSum)     'CheckSum 2 ~ 28 까지 하위 1 Byte 값
    
    'For i = 0 To 29
    '    Debug.Print i & " : " & Hex(SND_KOCOM(i))
    'Next i


End Sub
