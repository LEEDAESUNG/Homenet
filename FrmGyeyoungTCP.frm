VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmGyeyoungTCP 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "°è¿µÅë½Å(TCP) ¼¼´ëÅëº¸"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   14235
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.Frame Frame1 
      Caption         =   " ¼¼´ëÅëº¸ Å×½ºÆ® "
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   8850
      TabIndex        =   12
      Top             =   570
      Width           =   5325
      Begin VB.TextBox txt_Dong 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Text            =   "0113"
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txt_Ho 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Text            =   "0301"
         Top             =   330
         Width           =   915
      End
      Begin VB.CommandButton cmd_Test 
         Caption         =   "¼¼´ëÅëº¸ Å×½ºÆ®"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3210
         TabIndex        =   13
         Top             =   300
         Width           =   1875
      End
      Begin VB.Label Label2 
         Caption         =   "µ¿"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
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
         TabIndex        =   17
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "È£"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
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
         TabIndex        =   16
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.TextBox txt_HomeNet_IP 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1710
      TabIndex        =   5
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1755
   End
   Begin VB.TextBox txt_HomeNet_Port 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1710
      TabIndex        =   4
      Text            =   "12121"
      Top             =   480
      Width           =   1755
   End
   Begin VB.TextBox txt_HostPort 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1710
      TabIndex        =   3
      Text            =   "12121"
      Top             =   990
      Width           =   1755
   End
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "Á¾ ·á"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13080
      TabIndex        =   2
      Top             =   90
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
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
      TabIndex        =   1
      Top             =   1650
      Width           =   14160
   End
   Begin VB.CommandButton cmd_Clear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13050
      TabIndex        =   0
      Top             =   6390
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   990
      Top             =   6300
   End
   Begin MSWinsockLib.Winsock HomeSock 
      Left            =   540
      Top             =   6300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock HostSock 
      Left            =   90
      Top             =   6300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label lbl_GyeyoungTCPCnt 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5460
      TabIndex        =   11
      Top             =   150
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "GyeyoungTCP CNT : "
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
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
      TabIndex        =   10
      Top             =   150
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "HomeNet IP :"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Top             =   150
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "HomeNet Port :"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   8
      Top             =   510
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Host RCV Port :"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
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
      TabIndex        =   7
      Top             =   1020
      Width           =   1515
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   30
      X2              =   14190
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Label lbl_Date 
      Caption         =   "Timer"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9780
      TabIndex        =   6
      Top             =   180
      Width           =   3075
   End
End
Attribute VB_Name = "FrmGyeyoungTCP"
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
    Call GyeyoungTCP_Proc(txt_Dong, txt_Ho, "1234")
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

Call HomeLogger("[HomeNet Program ] GyeyoungTCP HomeNet Start..!!")

Timer1.Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim msg, Style, Title, Response
Dim Ret As Boolean
msg = " È¨³Ý ÇÁ·Î±×·¥À» Á¾·áÇÏ½Ã°Ú½À´Ï±î?         "
Style = vbYesNo + vbCritical + vbDefaultButton2
Title = " Parking Manager¢â "
Response = MsgBox(msg, Style, Title)
If Response = vbYes Then
    Call HomeLogger("[HOME Exit Proc]    " & " HomeNet Program END..!!")
    'Call HomeDB_Close(adoHome)
    End
End If
Me.MousePointer = 0
Cancel = True
End Sub


Private Sub HostSock_Close()
Call HomeLogger("[HostServer]    " & " Closed")
End Sub

'HOST·ÎºÎÅÍ Home_UDP ¹Þ±â
Private Sub HostSock_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Dim Tmp_Path As String
    Dim i, GateNo As Integer
    Dim CarNum As String
    Dim Dong As Integer
    Dim Ho As Integer
    Dim CarNo As String
    
    
On Error GoTo Err_P
    
    HostSock.GetData sData, , bytesTotal
    Call HomeLogger("HostSock UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sData)
    
    sData = Trim(sData)
    Dong = Val(Left(sData, 4))
    Ho = Val(Mid(sData, 5, 4))
    CarNo = Right(sData, 4)
    
    Call GyeyoungTCP_Proc(Dong, Ho, CarNo)

Exit Sub

Err_P:
    Call HomeLogger(" [HostSock UDP DataArrival]  " & Err.Description)

End Sub

Private Sub HomeSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'On Error GoTo Err_P
    'HomeNet.Close
    Call HomeLogger(" [HomeSock TCP Error]  " & Description)
'Err_P:
'    Call HomeLogger(" [HomeSock TCP Error]  " & Description)
End Sub




Private Sub GyeyoungTCP_Proc(Dong As Integer, Ho As Integer, CarNo As String)

    Dim i, s As Integer
    
On Error GoTo Err_Proc

    Call GetHeader(Dong, Ho, CarNo)
    
    Call HomeLogger("[GyeyoungTCP_Proc]  SND : " & CarNo & " " & Dong & "µ¿ " & Ho)
    
    Call Socket_Connect
    
    Exit Sub

Err_Proc:
    Call HomeLogger("[GyeyoungTCP_Proc]  " & Err.Description)
End Sub


Private Sub Socket_Connect()

On Error GoTo Err_P

    Call HomeLogger("[HomeAlarm_SocketConnect] È¨³ÝÁ¢¼Ó½Ãµµ : " & HomeNet_IP & " " & HomeNet_Port)
    If (HomeSock.State <> sckClosed) Then
        HomeSock.Close
    End If
    
    HomeSock.Connect HomeNet_IP, HomeNet_Port
        
Exit Sub
    
Err_P:
        Call HomeLogger("[HomeAlarm_Socket_Connect] ¿¡·¯³»¿ë : " & Err.Description)

End Sub

Private Sub HomeSock_Connect()
Dim sData As String
Dim bdata() As Byte

    'sdata = Socket_Data

    'ReDim bData(Len(sdata) - 1) As Byte
    'bData = StrConv(sdata, vbNarrow)
    'bData = StrConv(sdata, vbUnicode) 'ÇÑ±Û±úÁü
    'HomeSock.SendData bData
    'HomeSock.SendData sdata
    HomeSock.SendData GY_Tcp_Header

Exit Sub

Err_P:
    Call HomeLogger(" [HomeSock_Connect] Error : " & Err.Description)

End Sub

Private Sub HomeSock_DataArrival(ByVal bytesTotal As Long)

    Dim rMsg As String
    Dim B() As Byte
    Dim Ret As Integer
    Dim i As Integer
    Dim sData As String

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
    
    Call HomeLogger("[GyeyoungTCP_Proc]  RCV : " & rMsg)
    
    HomeSock.Close
    
Exit Sub

Err_P:
    Call HomeLogger(" [HomeSock_DataArrival RCV] ¿¡·¯³»¿ë : " & Err.Description)
End Sub

Private Sub Timer1_Timer()
    lbl_Date.Caption = Format(Now, "YYYY-MM-DD HH:NN:SS")
    
    If Format(Now, "NNSS") = "0001" Then
        List1.Clear
    End If
End Sub


Private Sub GetHeader(Dong As Integer, Ho As Integer, CarNum As String)

    ReDim GY_Tcp_Header(41) As Byte
    Dim DataLen, Data, DongHo As String
    Dim sDong As String
    Dim sHo As String
    Dim i, s As Integer
    
On Error GoTo Err_P

    'CarNum = Right(CarNo, 4)
    'Data = "2008-09-18,10:15:10,0,1234,12345678"
    Data = Format(Now, "YYYY-MM-DD") & "," & Format(Now, "HH:NN:DD") & ",0," & CarNum & "," & "12345678"

    sDong = Hex(Dong)
    sHo = Hex(Ho)
    DataLen = Hex(Len(Data))

    Select Case Len(sDong)
    Case 2
        GY_Tcp_Header(0) = &H0
        GY_Tcp_Header(1) = "&H" & sDong
    
    Case 3
        GY_Tcp_Header(0) = "&H" & Left(sDong, 1)
        GY_Tcp_Header(1) = "&H" & Right(sDong, 2)

    Case 4
        GY_Tcp_Header(0) = "&H" & Left(sDong, 2)
        GY_Tcp_Header(1) = "&H" & Right(sDong, 2)
    End Select
    
    Select Case Len(sHo)
        Case 2
            GY_Tcp_Header(2) = &H0
            GY_Tcp_Header(3) = "&H" & sHo
        
        Case 3
            GY_Tcp_Header(2) = "&H" & Left(sHo, 1)
            GY_Tcp_Header(3) = "&H" & Right(sHo, 2)
    
        Case 4
            GY_Tcp_Header(2) = "&H" & Left(sHo, 2)
            GY_Tcp_Header(3) = "&H" & Right(sHo, 2)
    End Select
    
    GY_Tcp_Header(4) = &H1
    
    Select Case Len(DataLen)
        Case 2
            GY_Tcp_Header(5) = &H0
            GY_Tcp_Header(6) = "&H" & DataLen
        
        Case 3
            GY_Tcp_Header(5) = "&H" & Left(DataLen, 1)
            GY_Tcp_Header(6) = "&H" & Right(DataLen, 2)
    
        Case 4
            GY_Tcp_Header(5) = "&H" & Left(DataLen, 2)
            GY_Tcp_Header(6) = "&H" & Right(DataLen, 2)
    End Select
    
    'For i = 0 To 6
    '    Debug.Print Hex(GY_Tcp_Header(i))
    'Next i

    s = 1
    For i = 7 To 41
        GY_Tcp_Header(i) = "&H" & Hex(Asc(Mid(Data, s, 1)))
        'Debug.Print Hex(GY_Tcp_Header(i))
        s = s + 1
    Next i
    
    Exit Sub
    
Err_P:
    Call HomeLogger(" [GetHeader]  " & Err.Description & "       " & Dong & "µ¿ " & Ho & "È£ " & CarNo & "  Â÷·®")
End Sub

