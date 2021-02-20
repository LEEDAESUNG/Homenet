VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmGyeyoung 
   BorderStyle     =   1  '단일 고정
   Caption         =   "계영통신 세대통보"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   14610
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
      Left            =   9225
      TabIndex        =   10
      Top             =   5370
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
         Text            =   "1001"
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
   Begin VB.TextBox TxtOpen 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1245
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5340
      Width           =   1425
   End
   Begin VB.TextBox TxtStop 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1245
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5670
      Width           =   1425
   End
   Begin VB.TextBox TxtDuration 
      Height          =   315
      Left            =   1245
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   6000
      Width           =   1425
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
      Left            =   30
      TabIndex        =   0
      Top             =   570
      Width           =   14550
   End
   Begin MSWinsockLib.Winsock HostSock 
      Left            =   8490
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "Start Time : "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   90
      TabIndex        =   9
      Top             =   5370
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "End Time  : "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   90
      TabIndex        =   8
      Top             =   5700
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "Proc. Time : "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   6045
      Width           =   1185
   End
   Begin VB.Label lbl_HomePort 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "UDP Port :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11010
      TabIndex        =   3
      Top             =   240
      Width           =   3315
   End
   Begin VB.Label lbl_HomeState 
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   210
      Width           =   5145
   End
   Begin VB.Label Label1 
      Caption         =   "세대통보 서버 상태 : "
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
      Left            =   240
      TabIndex        =   1
      Top             =   150
      Width           =   1785
   End
End
Attribute VB_Name = "FrmGyeyoung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Test_Click()
    Dim Qry As String
    Dim sTestCarNo As String
    sTestCarNo = "01가1234"
    
    If Len(txt_Dong.Text) <> 0 And Len(txt_Ho.Text) <> 0 Then
        Qry = "Insert Into rxinout Values (1, 0, '" & Format(Now, "YYYY-MM-DD") & "', '" & Format(Now, "HH:NN:SS") & "', '" & Val(Trim(txt_Dong.Text)) & "-" & Val(Trim(txt_Ho.Text)) & "', '" & Trim(sTestCarNo) & "', '" & Trim(sTestCarNo) & "', 0, 0)"
        Call GyeYoung_Proc3(Qry)
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

If HomeDB_Open(adoHome) Then
    Call HomeLogger(" [Gyeyoung Program] Gyeyoung HomeNet Start..!!")
    'Call HomeDB_Close(adoHome)
Else
    Call HomeLogger(" [Gyeyoung Program] Gyeyoung HomeNet DB Connect Fail...!!")
End If

HostSock.Protocol = sckUDPProtocol
HostSock.LocalPort = HostPort
HostSock.Bind
FrmGyeyoung.lbl_HomePort.Caption = "UDP Port : " & HostPort

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim msg, Style, Title, Response
Dim Ret As Boolean
msg = " 세대통보 프로그램을 종료하시겠습니까?         "
Style = vbYesNo + vbCritical + vbDefaultButton2
Title = " Parking Manager™ "
Response = MsgBox(msg, Style, Title)
If Response = vbYes Then
    Call HomeLogger("[HOME Exit Proc]    " & " 세대통보 프로그램 종료")
    Call HomeDB_Close(adoHome)
    End
End If
Me.MousePointer = 0
Cancel = True
End Sub

'HOST로부터 Home_UDP 받기
Private Sub HostSock_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Dim Tmp_Path As String
    Dim I, GateNo As Integer
    Dim CarNum As String
    Dim Qry As String
    
On Error GoTo Err_P
    
    HostSock.GetData sData, , bytesTotal
    Call HomeLogger(" [HostSock UDP Port] " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sData)
    Call HomeNet_Proc(sData)
    
    HomeNet_CarNo = Mid(sData, 9)

    '수정시작
    '기존 '123456' 를 RF Code 값으로 대체(차량관제프로그램으로부터 HomeNet_CarNo 변수에 RF Code를 수신한다)
    'Qry = "Insert Into rxinout Values (1, 0, '" & Format(Now, "YYYY-MM-DD") & "', '" & Format(Now, "HH:NN:SS") & "', '" & Val(HomeNet_Dong) & "-" & Val(HomeNet_Ho) & "', '" & Right(Trim(HomeNet_CarNo), 4) & "', '123456', 0, 0)"
    'Qry = "Insert Into rxinout Values (1, 0, '" & Format(Now, "YYYY-MM-DD") & "', '" & Format(Now, "HH:NN:SS") & "', '" & Val(HomeNet_Dong) & "-" & Val(HomeNet_Ho) & "', '" & Right(Trim(sdata), 16) & "', '" & Right(Trim(sdata), 16) & "', 0, 0)"
    Qry = "Insert Into rxinout Values (1, 0, '" & Format(Now, "YYYY-MM-DD") & "', '" & Format(Now, "HH:NN:SS") & "', '" & Val(HomeNet_Dong) & "-" & Val(HomeNet_Ho) & "', '" & Trim(HomeNet_CarNo) & "', '" & Trim(HomeNet_CarNo) & "', 0, 0)"
    '수정끝
    Call HomeLogger(" [Gyeyoung Proc] " & Qry)
    Call GyeYoung_Proc3(Qry)

Exit Sub

Err_P:
    Call HomeLogger(" [HostSock UDP DataArrival]  " & Err.Description)

End Sub

Public Sub HostSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call HomeLogger(" [HostSock UDP Error]  " & Description)
End Sub


Public Sub GyeYoung_Proc2(Qry As String)
Dim sngStart    As Single
Dim sngStop     As Single
Dim sngDuration As Single
    
    Me.MousePointer = 11
    sngStart = Timer
    TxtOpen = CStr(sngStart): TxtOpen.Refresh
    
    
    If HomeDB_Open(adoHome) Then
        adoHome.Execute Qry
        
        sngStop = Timer
        TxtStop = CStr(sngStop): TxtStop.Refresh
        TxtDuration = sngStop - sngStart
        TxtDuration.Refresh
        
        Call HomeLogger(" [Gyeyoung_Proc] 세대통보 성공..!!  Lap Time : " & TxtDuration & " Sec")
        Call HomeLogger(" [Gyeyoung_Proc] -----------------------------------------------------")
        Call HomeDB_Close(adoHome)
    Else
        Call HomeLogger(" Home DB Reconnection Fail..!!")
        sngStop = Timer
        TxtStop = CStr(sngStop): TxtStop.Refresh
        TxtDuration = sngStop - sngStart
        TxtDuration.Refresh
        Me.MousePointer = 0
        Exit Sub
    End If

Me.MousePointer = 0

Exit Sub

Err_Proc:
    Me.MousePointer = 0
    
    Call HomeLogger("[GyeYoung_Proc]  CarNo : " & CarNo & " / 동 - 호 : " & Dong & " - " & Ho & " / Error Msg : " & Err.Description)
    Call DataBaseClose(adoHome)
    sngStop = Timer
    TxtStop = CStr(sngStop): TxtStop.Refresh
    TxtDuration = sngStop - sngStart
    TxtDuration.Refresh
    
End Sub

Public Sub GyeYoung_Proc3(Qry As String)
Dim sngStart    As Single
Dim sngStop     As Single
Dim sngDuration As Single
    
On Error GoTo Err_Proc
    
Me.MousePointer = 11
        adoHome.Execute "Show Variables"
        
Return_GY:
        
        
        sngStart = Timer
        TxtOpen = CStr(sngStart): TxtOpen.Refresh
    
        adoHome.Execute Qry
        
        sngStop = Timer
        TxtStop = CStr(sngStop): TxtStop.Refresh
        TxtDuration = sngStop - sngStart
        TxtDuration.Refresh
        
        Call HomeLogger(" [Gyeyoung_Proc] 세대통보 성공..!!  Lap Time : " & TxtDuration & " Sec")
        Call HomeLogger(" [Gyeyoung_Proc] -----------------------------------------------------")

Me.MousePointer = 0

Exit Sub

Err_Proc:
    Me.MousePointer = 0
    Call HomeLogger("[Gyeyoung_Proc]  Error Msg : " & Err.Description)

    Call HomeDB_Close(adoHome)
    If HomeDB_Open(adoHome) Then
        Me.MousePointer = 11
        GoTo Return_GY
    Else
        Call HomeLogger("[Gyeyoung_Proc]  DB ReConnection Fail..!!  " & Qry)
        sngStop = Timer
        TxtStop = CStr(sngStop): TxtStop.Refresh
        TxtDuration = sngStop - sngStart
        TxtDuration.Refresh
    End If
    
End Sub

'Public Sub GyeYoung_Proc(Dong As Integer, Ho As Integer, CarNo As String)
'Dim Qry As String
'
'    Qry = "Insert Into rxinout Values (1, 0, '" & Format(Now, "YYYY-MM-DD") & "', '" & Format(Now, "HH:NN:SS") & "', '" & Dong & "-" & Ho & "', '" & Right(CarNo, 4) & "', '123456', 0, 0)"
'    'Qry = "Insert Into rxinout Values (1, 0, '" & Format(Now, "YYYY-MM-DD") & "', '" & Format(Now, "HH:NN:SS") & "', '" & Dong & "-" & Ho & "', '" & Right(CarNo, 4) & "', '123456', 0, 0, Null)"
'    FrmTcpServer.Home_sock.SendData (Qry)
'    Call DataLogger("[Home UDP 전송]  DATA = " & Qry)
'    Call DataLogger("[GyeYoung_Proc]  CarNo : " & CarNo & " / 동 - 호 : " & Dong & " - " & Ho & " / 세대통보")
'Exit Sub
'
'Err_Proc:
'    Call DataLogger("[GyeYoung_Proc]  CarNo : " & CarNo & " / 동 - 호 : " & Dong & " - " & Ho & " / Error Msg : " & Err.Description)
'End Sub
Private Sub lbl_HomeState_Click()

End Sub
