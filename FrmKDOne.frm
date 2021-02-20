VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmKDOne 
   Caption         =   "KD One"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   17190
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
      Left            =   11730
      TabIndex        =   10
      Top             =   570
      Width           =   5325
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
         TabIndex        =   12
         Text            =   "0202"
         Top             =   330
         Width           =   915
      End
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
         TabIndex        =   11
         Text            =   "0101"
         Top             =   330
         Width           =   915
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
         TabIndex        =   15
         Top             =   360
         Width           =   315
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
         Index           =   0
         Left            =   1260
         TabIndex        =   14
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
      Left            =   1680
      TabIndex        =   5
      Text            =   "127.0.0.1"
      Top             =   150
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
      Left            =   1680
      TabIndex        =   4
      Text            =   "12121"
      Top             =   510
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
      Left            =   1680
      TabIndex        =   3
      Text            =   "12121"
      Top             =   1020
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
      Left            =   15930
      TabIndex        =   2
      Top             =   120
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
      ItemData        =   "FrmKDOne.frx":0000
      Left            =   30
      List            =   "FrmKDOne.frx":0002
      TabIndex        =   1
      Top             =   1680
      Width           =   17160
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
      Left            =   15990
      TabIndex        =   0
      Top             =   6420
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   6330
   End
   Begin MSWinsockLib.Winsock HomeSock 
      Left            =   510
      Top             =   6330
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock HostSock 
      Left            =   60
      Top             =   6330
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
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
      Left            =   150
      TabIndex        =   9
      Top             =   180
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
      Left            =   150
      TabIndex        =   8
      Top             =   540
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
      Left            =   150
      TabIndex        =   7
      Top             =   1050
      Width           =   1515
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   17190
      Y1              =   1560
      Y2              =   1560
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
      Left            =   9750
      TabIndex        =   6
      Top             =   210
      Width           =   3075
   End
End
Attribute VB_Name = "FrmKDOne"
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
        Call KDOne_Proc(Trim(txt_Dong.Text), Trim(txt_Ho.Text), "01°¡1234")
    Else
        MsgBox ("Å×½ºÆ®ÇÒ µ¿/È£¸¦ È®ÀÎÇÏ¼¼¿ä..!!")
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

Call HomeLogger("[HomeNet Program ] KD One HomeNet Start..!!")

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

'HOST·ÎºÎÅÍ Home_UDP ¹Þ±â
Private Sub HostSock_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Dim Tmp_Path As String
    Dim i, GateNo As Integer
    Dim CarNum As String
    
On Error GoTo Err_P
    
    HostSock.GetData sData, , bytesTotal
    Call HomeLogger("HostSock UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sData)
    
    Call HomeNet_Proc(sData)
    
    Call KDOne_Proc(HomeNet_Dong, HomeNet_Ho, HomeNet_CarNo)

Exit Sub

Err_P:
    Call HomeLogger(" [HostSock UDP DataArrival]  " & Err.Description)

End Sub

Private Sub HomeSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call HomeLogger(" [HomeSock UDP Error]  " & Description)
End Sub

Private Sub Timer1_Timer()
lbl_Date.Caption = Format(Now, "YYYY-MM-DD HH:NN:SS")

If Format(Now, "NNSS") = "0001" Then
    List1.Clear
End If

End Sub


Public Sub KDOne_Proc(Dong As String, Ho As String, CarNo As String)
Dim i, q, r As Integer
Dim s, t As String
Dim bdata() As Byte
Dim xData() As Byte

On Error GoTo Err_Proc

'CarNo = "¼­¿ï01°¡2345"
'
'ReDim bData(LenH(CarNo))
'bData = Utf8BytesFromString(CarNo)
'
'For i = 0 To (LenH(CarNo) - 1)
'    Debug.Print bData(i)
'Next i

CarNo = Trim(CarNo)
CarNo = Right(CarNo, 4)

Socket_Data = ""
Socket_Data = "<HNML>" & _
                "<ControlRequest TransID=" & Chr(34) & "CAR" & Format(Now, "YYYYMMDDHHNNSS") & Chr(34) & ">" & _
                "<FunctionID>1F03010B</FunctionID>" & _
                "<FunctionCategory>Control</FunctionCategory>" & _
                "<InputList size=" & Chr(34) & "1" & Chr(34) & ">" & _
                "<Input size=" & Chr(34) & "6" & Chr(34) & ">" & _
                "<Data name=" & Chr(34) & "Complex" & Chr(34) & ">0000</Data>" & _
                "<Data name=" & Chr(34) & "Dong" & Chr(34) & ">" & Format(Dong, "0000") & "</Data>" & _
                "<Data name=" & Chr(34) & "Ho" & Chr(34) & ">" & Format(Ho, "0000") & "</Data>" & _
                "<Data name=" & Chr(34) & "CarNo" & Chr(34) & ">" & CarNo & "</Data>" & _
                "<Data name=" & Chr(34) & "Direction" & Chr(34) & ">In</Data>" & _
                "<Data name=" & Chr(34) & "Time" & Chr(34) & ">" & Format(Now, "YYYY-MM-DD") & "T" & Format(Now, "HH:NN:SS") & "</Data>" & _
                "</Input></InputList></ControlRequest>" & _
                "</HNML>"

i = 100 + LenH(Socket_Data)
r = i
s = Hex(i)

ReDim KD_Header(r - 1) As Byte
bdata = StrConv(Socket_Data, vbFromUnicode)
q = 100
For i = 0 To (LenH(Socket_Data) - 1)
    KD_Header(q) = bdata(i)
    q = q + 1
Next i
'Debug.Print KD_Header(r - 1)

'Exit Sub

'MessageStart
KD_Header(0) = &HA0
KD_Header(1) = &H0
'Version
KD_Header(2) = &H1
'Flag
KD_Header(3) = &H0

'Length
KD_Header(4) = &H0
KD_Header(5) = &H0
KD_Header(6) = &H1
KD_Header(7) = "&H" & Mid(s, 2, 2)

'MessageID
KD_Header(8) = &H0
KD_Header(9) = &H0
KD_Header(10) = &H0
KD_Header(11) = &H0

'Sequence Number
KD_Header(12) = &H0
KD_Header(13) = &H0
KD_Header(14) = &H0
KD_Header(15) = &H0

'Sourece ID
t = "00000000000000001701"
ReDim bdata(19) As Byte
bdata = Utf8BytesFromString(t)
q = 16
For i = 0 To 19
    KD_Header(q) = bdata(i)
    q = q + 1
Next i
''´ÜÁö 4 byte
'KD_Header(16) = &H0
'KD_Header(17) = &H0
'KD_Header(18) = &H0
'KD_Header(19) = &H0
''µ¿ 4 byte
'KD_Header(20) = &H0
'KD_Header(21) = &H0
'KD_Header(22) = &H0
'KD_Header(23) = &H0
''È£ 4 byte
'KD_Header(24) = &H0
'KD_Header(25) = &H0
'KD_Header(26) = &H0
'KD_Header(27) = &H0
''È®Àå 2 byte
'KD_Header(28) = &H0
'KD_Header(29) = &H0
''±â±âID 6 byte
'KD_Header(30) = &H0
'KD_Header(31) = &H0
'KD_Header(32) = &H1
'KD_Header(33) = &H7
'KD_Header(34) = &H0
'KD_Header(35) = &H1


'Destination ID
t = "00000000000000001F01"
ReDim bdata(19) As Byte
bdata = Utf8BytesFromString(t)
q = 36
For i = 0 To 19
    KD_Header(q) = bdata(i)
    q = q + 1
Next i

''´ÜÁö 4 byte
'KD_Header(36) = &H0
'KD_Header(37) = &H0
'KD_Header(38) = &H0
'KD_Header(39) = &H0
''µ¿ 4 byte
'KD_Header(40) = &H0
'KD_Header(41) = &H0
'KD_Header(42) = &H0
'KD_Header(43) = &H0
''È£ 4 byte
'KD_Header(44) = &H0
'KD_Header(45) = &H0
'KD_Header(46) = &H0
'KD_Header(47) = &H0
''È®Àå 2 byte
'KD_Header(48) = &H0
'KD_Header(49) = &H0
''±â±âID 6 byte
'KD_Header(50) = &H0
'KD_Header(51) = &H0
'KD_Header(52) = &H1
'KD_Header(53) = &HF
'KD_Header(54) = &H0
'KD_Header(55) = &H2

'OP Code
KD_Header(56) = &H2F
KD_Header(57) = &HF1

'TransactionID

t = Format(Now, "YYYY-MM-DD") & "T" & Format(Now, "HH:NN:SS") & " "
ReDim bdata(Len(t) - 1) As Byte
'bData = Utf8BytesFromString(t)
bdata = StrConv(t, vbFromUnicode)

q = 58
For i = 0 To 19
    KD_Header(q) = bdata(i)
    q = q + 1
Next i

'CRC
KD_Header(78) = &H0
KD_Header(79) = &H0
KD_Header(80) = &H0
KD_Header(81) = &H0

'Optional Header / DATA
KD_Header(82) = &H0
KD_Header(83) = &H0
KD_Header(84) = &H0
KD_Header(85) = &H0
KD_Header(86) = &H0
KD_Header(87) = &H0
KD_Header(88) = &H0
KD_Header(89) = &H0
KD_Header(90) = &H0
KD_Header(91) = &H0
KD_Header(92) = &H0
KD_Header(93) = &H0
KD_Header(94) = &H0
KD_Header(95) = &H0
KD_Header(96) = &H0
KD_Header(97) = &H0

'MessageEnd
KD_Header(98) = &H0
KD_Header(99) = &HA

'For i = 0 To (r - 1)
'    Call HomeLogger(i & " : " & KD_Header(i) & "    " & Chr(KD_Header(i)))
'Next i

Call HomeLogger("[KDOne_Proc]  SND : " & Socket_Data)
'Debug.Print Socket_Data
Call Socket_Connect

Exit Sub

Err_Proc:
    Call HomeLogger("[KDOne_Proc]  " & Err.Description)
End Sub

Public Sub Socket_Connect()
Dim bdata() As Byte

On Error GoTo Err_P
    Call HomeLogger("[HomeAlarm_SocketConnect] È¨³ÝÁ¢¼Ó½Ãµµ : " & HomeNet_IP & " " & HomeNet_Port)
    If (HomeSock.State <> sckClosed) Then
        HomeSock.Close
    End If
    HomeSock.Connect HomeNet_IP, HomeNet_Port

Exit Sub

Err_P:
    Call HomeLogger("[HomeAlarm_SocketConnect] ¿¡·¯³»¿ë : " & Err.Description)

End Sub

Private Sub HomeSock_Connect()
Dim sData As String
Dim bdata() As Byte
Dim i As Integer

On Error GoTo Err_P

    'sdata = Trim(Socket_Data)
    'ReDim bData(Len(sdata) - 1) As Byte
    
    'bData = StrConv(sdata, vbUnicode)
    'bData = Utf8BytesFromString(sdata)
     
    HomeSock.SendData KD_Header
    
    
    
    'HomeSock.SendData bData
    
'    Debug.Print bData(0)
'    Debug.Print bData(1)
        
    Socket_Data = ""

Exit Sub

Err_P:
    Call HomeLogger(Format(Now, "yyyy-mm-dd hh:nn:ss") & " [HomeSock_Connect] Error : " & Err.Description)

End Sub

Private Sub HomeSock_DataArrival(ByVal bytesTotal As Long)
Dim rMsg As String
Dim B() As Byte
Dim Ret As Integer
Dim i As Integer
Dim sData As String

On Error GoTo Err_P

HomeSock.GetData sData, , bytesTotal
Call HomeLogger("HostSock UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sData)
Call HomeLogger("[KDOne_Proc]  RCV : " & Mid(sData, 16, (FormatNumber(bytesTotal, 0, , , vbTrue)) - 16))

HomeSock.Close

Exit Sub

Err_P:
    Call HomeLogger(" [KDOne_Proc RCV] ¿¡·¯³»¿ë : " & Err.Description)
End Sub



