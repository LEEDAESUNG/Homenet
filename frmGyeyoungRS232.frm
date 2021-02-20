VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmGyeyoungRS232 
   BorderStyle     =   1  '단일 고정
   Caption         =   "계영통신(시리얼) 세대통보"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14220
   Icon            =   "frmGyeyoungRS232.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   14220
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CheckBox chk_Debug 
      Caption         =   "Alive Check"
      Height          =   375
      Left            =   11430
      TabIndex        =   14
      Top             =   6360
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   1695
      TabIndex        =   12
      Text            =   "12121"
      Top             =   525
      Width           =   1755
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
      Left            =   8850
      TabIndex        =   6
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   360
         Width           =   315
      End
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
      ItemData        =   "frmGyeyoungRS232.frx":A5A2
      Left            =   30
      List            =   "frmGyeyoungRS232.frx":A5A4
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
   Begin VB.Label Label1 
      Caption         =   "HomeNet Serial:"
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
      Left            =   165
      TabIndex        =   13
      Top             =   555
      Width           =   1515
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   150
      Width           =   1515
   End
End
Attribute VB_Name = "frmGyeyoungRS232"
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
    'Call GyeyoungRS232_Alarm(0, "서울01가9012", "1234", "5678")
    Call GyeyoungRS232_Alarm(0, "1234", txt_Dong, txt_Ho)
    Call HomeLogger(" [HomeAlarm Proc]  Sending... " & "1234" & " " & txt_Dong & "동 " & txt_Ho & "동 ")
    Call HomeLogger(" [KDOne_Proc]  SND : <HNML><ControlRequest TransID=-CAR20161020132845-><FunctionID>1F03010B</FunctionID><FunctionCategory>Control</FunctionCategory><InputList size=-1-><Input size=-6-><Data name=-Complex->0000</Data><Data name=-Dong->0101</Data><Data name=-Ho->0202</Data><Data name=-CarNo->5678</Data><Data name=-Direction->In</Data><Data name=-Time->2016-10-20T13:28:45</Data></Input></InputList></ControlRequest></HNML>")
End Sub

Private Sub Form_Load()

    If App.PrevInstance = True Then
        End
    End If
    
    On Error GoTo Err_P

    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    
    IniFileName$ = App.Path & "\HomeNet.ini"
    Doc_Path_Name$ = App.Path & "\Doc\"

    txt_HostPort.Text = HostPort
    
    HostSock.Protocol = sckUDPProtocol
    HostSock.LocalPort = HostPort
    HostSock.Bind
    
    Call HomeLogger("[HomeNet Program] Gyeyoung Serial HomeNet Start..!!")
    
    '세대통보 시리얼포트 열기 (서울통신)
    MSComm1.CommPort = HomeNet_ComPort
    MSComm1.Settings = "9600,e,8,1"
    MSComm1.InputLen = 1
    MSComm1.InputMode = comInputModeBinary
    MSComm1.PortOpen = True
    
    If (MSComm1.PortOpen = True) Then
        Call HomeLogger("[HomeNet Proc] ComPort Open Success..!!")
        txt_HomeNet_Port = "Com" & MSComm1.CommPort
    Else
        Call HomeLogger("[HomeNet Proc] ComPort Open Failure..!!")
        txt_HomeNet_Port = "Com" & MSComm1.CommPort
        MsgBox "    세대통보 시리얼포트 오픈 실패..!!"
        End
    End If
    
    Timer1.Enabled = True

    Exit Sub
Err_P:
    txt_HomeNet_Port = "Com" & HomeNet_ComPort
    Call HomeLogger("[HomeNet ComPort Proc] " & Err.Description)
    
End Sub

Private Sub MSComm1_OnComm()
    Dim i, Cnt As Integer
    Dim tmpinstring() As Byte
    Dim Instring As String
    Dim Data(3) As Byte
    Dim bchk As Byte
    Dim Index As Integer
    
'On Error GoTo Err_Proc

Select Case MSComm1.CommEvent
           Case comEvDSR
           Case comEvSend
           Case comEvCTS
           Case comEvReceive
                '한바이트를 읽어온다
                tmpinstring = MSComm1.Input
                Instring = Format$(Hex(tmpinstring(i)), "00")
                If Instring = "E8" Then
                    GbDATA = True
                End If
                
                If GbDATA = True Then
                    GsData = GsData & Instring
                End If

                
                If Len(GsData) = 10 Then

                    If Mid$(GsData, 1, 4) <> "E8E8" Then
                        GsData = ""
                        GbDATA = False
                        MSComm1.InBufferCount = 0

                        Exit Sub
                    End If



                    Select Case Mid$(GsData, 7, 2)
                        Case "71"   'Intial Data
                            'Call HomeLogger(" [HomeAlarm Proc] Code 71 Initial Data RCV..!!")
                            Call None_Delay_Time(0.4)
                            Data(0) = &HBB
                            Data(1) = &HA4
                            Data(2) = &H71
                            bchk = Data(0)
                            For i = 1 To 2
                                bchk = bchk Xor Data(i)
                            Next
                            Data(3) = bchk
                            MSComm1.Output = Data

                            'Call HomeLogger(" [HomeAlarm Proc] Code 71 Initial Data 응답전송..!!")

                        Case "33"   'Status Data Free / Busy

                             If Glo_Event = True Then
                                Call None_Delay_Time(0.4)
                                MSComm1.Output = Glo_bData
                                'Call HomeLogger(" [HomeAlarm Proc] Car IN Event SND..!!")
                                'Call HomeLogger(" [HomeAlarm Proc] SND..!! 입차통보 중... : " & "(" & GsData & ")")
                                'Glo_Event = False
                             Else
                                Call None_Delay_Time(0.4)
                                Data(0) = &HBB
                                Data(1) = &HA4
                                Data(2) = &H33
                                bchk = Data(0)
                                For i = 1 To 2
                                    bchk = bchk Xor Data(i)
                                Next
                                Data(3) = bchk
                                MSComm1.Output = Data
'                                If chk_Debug.value = 1 Then
'                                    Call HomeLogger(" [HomeAlarm Proc] Code 33 Alive Satus RCV..!!")
'                                End If
                             End If

                        Case "34"
                            Call None_Delay_Time(0.4)
                            Data(0) = &HBB
                            Data(1) = &HA4
                            Data(2) = &H34
                            bchk = Data(0)
                            For i = 1 To 2
                                bchk = bchk Xor Data(i)
                            Next
                            Data(3) = bchk
                            MSComm1.Output = Data
                            'Call HomeLogger(" [HomeAlarm Proc] RCV..!! Code : 34 : 데이터패킷 응답전송 ")
                        Case "51"
                            Glo_Event = False
                            'Call HomeLogger(" [HomeAlarm Proc] RCV..!! Code : 51 : 차량입차 ACK : " & GsData)
                            Call HomeLogger(" [HomeAlarm Proc] RCV..!! 입차통보 정상 ")
                        Case "52"
                            'Call HomeLogger(" [HomeAlarm Proc] RCV..!! Code : 52 : 입차통보 에러 : " & GsData)
                            Call HomeLogger(" [HomeAlarm Proc] RCV..!! 입차통보 에러 ")
                        Case "5C"
                            'Call HomeLogger(" [HomeAlarm Proc] RCV..!! Code : 5C : 차량출차 : " & GsData)
                            Call HomeLogger(" [HomeAlarm Proc] RCV..!! 출차통보 정상 ")
                        Case "5D"
                            'Call HomeLogger(" [HomeAlarm Proc] RCV..!! Code : 5D : 출차통보 에러 : " & GsData)
                            Call HomeLogger(" [HomeAlarm Proc] RCV..!! 출차통보 에러 ")
                        Case Else
                            Call HomeLogger(" [HomeAlarm Proc] RCV..!! Code : 정의되지 않은 패킷 에러 : " & GsData)
                    End Select
                   GsData = ""
                   GbDATA = False
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
        Timer1.Enabled = False
        Call HomeLogger(" [HomeAlarm END Proc] Program END..!!")
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
    
    '계영통신 세대통보
    CarNum = Trim(HomeNet_CarNo)
    CarNum = Right(CarNum, 4)
    Call HomeLogger(" [HomeAlarm Proc] SND.... " & CarNum & " " & HomeNet_Dong & "동 " & HomeNet_Ho & "동 ")
    Call GyeyoungRS232_Alarm(0, CarNum, HomeNet_Dong, HomeNet_Ho)

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
End Sub

Public Sub GyeyoungRS232_Alarm(inout As Integer, tmpCarNo As String, tmpDong As String, tmpHo As String)

    Dim i As Integer
    Dim DTime As String
    Dim SND_KOCOM(30) As Byte
    Dim CheckSum As Long
    'Dim tmpCarNo As String
    
    Glo_bData(1) = &HBB
    Glo_bData(2) = &HAE
    Glo_bData(3) = &H51

    Glo_bData(4) = "&H" & Hex(Mid$(Format(tmpDong, "0000"), 1, 2))    '&H01
    Glo_bData(5) = "&H" & Hex(Mid$(Format(tmpDong, "0000"), 3, 2))    '&H01

    Glo_bData(6) = "&H" & Hex(Mid$(Format(tmpHo, "0000"), 1, 2))
    Glo_bData(7) = "&H" & Hex(Mid$(Format(tmpHo, "0000"), 3, 2))

    Glo_bData(8) = "&H" & Hex(Mid$(tmpCarNo, 1, 2))
    Glo_bData(9) = "&H" & Hex(Mid$(tmpCarNo, 3, 2))

    Glo_bData(10) = "&H" & Hex(Format$(Now, "MM"))
    Glo_bData(11) = "&H" & Hex(Format$(Now, "DD"))
    Glo_bData(12) = "&H" & Hex(Format$(Now, "HH"))
    Glo_bData(13) = "&H" & Hex(Format$(Now, "NN"))

    bBCC = Glo_bData(1)
    For i = 2 To 13
        bBCC = bBCC Xor Glo_bData(i)
    Next
    Glo_bData(14) = bBCC
    Glo_Event = True

End Sub
