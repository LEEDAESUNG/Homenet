VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmLGElectron 
   Caption         =   "LGElectron"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   17190
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   6330
   End
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
      Left            =   30
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
Attribute VB_Name = "FrmLGElectron"
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
    Call LGElectron_Proc(txt_Dong, txt_Ho, "01°¡5678")
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

Call HomeLogger("[HomeNet Program ] LGElectron HomeNet Start..!!")

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
    Dim sdata As String
    Dim Tmp_Path As String
    Dim I, GateNo As Integer
    Dim CarNum As String
    
On Error GoTo Err_P
    
    HostSock.GetData sdata, , bytesTotal
    Call HomeLogger("HostSock UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sdata)
    
    Call HomeNet_Proc(sdata)
    
    Call LGElectron_Proc(HomeNet_Dong, HomeNet_Ho, HomeNet_CarNo)

Exit Sub

Err_P:
    Call HomeLogger(" [HostSock UDP DataArrival]  " & Err.Description)

End Sub

Private Sub HomeSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo Err_P
    'HomeNet.Close
Err_P:
    Call HomeLogger(" [HomeSock TCP Error]  " & Description)
End Sub




Public Sub LGElectron_Proc(Dong As String, Ho As String, CarNo As String)

On Error GoTo Err_Proc

    HomeNet_Dong = Trim(Dong)
    HomeNet_Ho = Trim(Ho)
    'HomeNet_CarNo = Trim(CarNo)

    'rn=1 : Àü¼ÛÈ½¼ö
    Socket_Data = "SLL" & "&" & "8" & "&" & "7" & "&" & "cmd=1" & "&" & "rn=1" & "&" & "args_cnt=4" & "&" & "id=" & Format(HomeNet_Dong, "00000") & Format(HomeNet_Ho, "00000") & "&" & "name=no" & "&" & "info=" & Trim(CarNo) & "&" & "time=" & Format(Now, "yyyy³â mm¿ù ddÀÏ hh½Ã nnºÐ") & ""
    
    Call HomeLogger("[LGElectron_Proc]  SND : " & Socket_Data)
    
    Call Socket_Connect
    
    Exit Sub

Err_Proc:
    Call HomeLogger("[LGElectron_Proc]  " & Err.Description)
End Sub


Public Sub Socket_Connect()

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
Dim sdata As String
Dim bdata() As Byte

    sdata = Socket_Data

    'ReDim bData(Len(sdata) - 1) As Byte
    'bData = StrConv(sdata, vbNarrow)
    'bData = StrConv(sdata, vbUnicode) 'ÇÑ±Û±úÁü
    'HomeSock.SendData bData
    
    HomeSock.SendData sdata

Exit Sub

Err_P:
    Call HomeLogger(Format(Now, "yyyy-mm-dd hh:nn:ss") & " [HomeSock_Connect] Error : " & Err.Description)

End Sub

Private Sub HomeSock_DataArrival(ByVal bytesTotal As Long)
Dim rMsg As String
    
    Dim sdata As String

    HomeSock.GetData sdata, , bytesTotal
    
    Call HomeLogger("[LGElectron_Proc]  RCV : " & sdata)

'    Ret = InStr(sdata, "res=ok")
'    If (Ret > 0) Then
'        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [´ÜÁö¼­¹öÆÐÅ¶ ¼ö½Å °á°ú] " & "Â÷·® ÁøÀÔ ¾Ë¸² ÀÀ´ä ¼º°ø"
'    End If
'
'    Ret = InStr(sdata, "res=fail")
'    If (Ret > 0) Then
'        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [´ÜÁö¼­¹öÆÐÅ¶ ¼ö½Å °á°ú] " & "Â÷·® ÁøÀÔ ¾Ë¸² ÀÀ´ä ½ÇÆÐ"
'    End If

    
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

