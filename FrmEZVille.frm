VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmEZVille 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "ÀÌÁöºô È¨³Ý ÇÁ·Î±×·¥"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   461
   ScaleMode       =   3  'ÇÈ¼¿
   ScaleWidth      =   1156
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
      Left            =   11880
      TabIndex        =   16
      Top             =   540
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
         TabIndex        =   19
         Text            =   "0101"
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
         TabIndex        =   18
         Text            =   "0202"
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
         TabIndex        =   17
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
         Index           =   0
         Left            =   1260
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.TextBox txt_ezVilleDong 
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
      Left            =   5250
      TabIndex        =   11
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txt_ezVilleHo 
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
      Left            =   5250
      TabIndex        =   10
      Text            =   "12121"
      Top             =   480
      Width           =   1095
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
      Left            =   1740
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
      Left            =   1740
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
      Left            =   1740
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
      Left            =   16080
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
      ItemData        =   "FrmEZVille.frx":0000
      Left            =   60
      List            =   "FrmEZVille.frx":0002
      TabIndex        =   1
      Top             =   1650
      Width           =   17250
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
      Left            =   16050
      TabIndex        =   0
      Top             =   6390
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1020
      Top             =   6300
   End
   Begin MSWinsockLib.Winsock HomeSock 
      Left            =   570
      Top             =   6300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock HostSock 
      Left            =   120
      Top             =   6300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label lbl_Cnt 
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
      Height          =   255
      Left            =   5250
      TabIndex        =   15
      Top             =   1050
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "ezVille CNT :"
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
      Index           =   5
      Left            =   3900
      TabIndex        =   14
      Top             =   1050
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "ezVille Dong :"
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
      Index           =   4
      Left            =   3900
      TabIndex        =   13
      Top             =   150
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "ezVille Ho :"
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
      Index           =   3
      Left            =   3900
      TabIndex        =   12
      Top             =   510
      Width           =   1515
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
      Left            =   210
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
      Left            =   210
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
      Left            =   210
      TabIndex        =   7
      Top             =   1020
      Width           =   1515
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   4
      X2              =   1150
      Y1              =   102
      Y2              =   102
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
      Left            =   12330
      TabIndex        =   6
      Top             =   180
      Width           =   3075
   End
End
Attribute VB_Name = "FrmEZVille"
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
        Call EasyVil_Alarm(0, "01°¡2345", Format(txt_Dong.Text, "0000"), Format(txt_Ho.Text, "0000"))
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

txt_ezVilleDong.Text = ezVille_Dong
txt_ezVilleHo.Text = ezVille_Ho

HostSock.Protocol = sckUDPProtocol
HostSock.LocalPort = HostPort
HostSock.Bind

Call HomeLogger(" [ezVille Program] ezVille HomeNet Start..!!")

'ÀÌÁöºô ¹ÙÀÎµù==============================================================================
EasyVil_Mode = "INIT"
Call Socket_Connect_EasyVil

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
    Dim sdata As String
    Dim Tmp_Path As String
    Dim i, GateNo As Integer
    Dim CarNum As String
    
On Error GoTo Err_P
    
    HostSock.GetData sdata, , bytesTotal
    Call HomeLogger(" [HostSock UDP Port] " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sdata)
    
    Call HomeNet_Proc(sdata)
    
    'EasyVil_Alarm(inout As Integer, CarNo As String, tmpDong As Integer, tmpHo As Integer)
    'Çö´ëÅë½Å ¼¼´ëÅëº¸
    Call EasyVil_Alarm(0, HomeNet_CarNo, HomeNet_Dong, HomeNet_Ho)

Exit Sub

Err_P:
    Call HomeLogger(" [HostSock UDP DataArrival]  " & Err.Description)

End Sub

Private Sub Timer1_Timer()
lbl_Date.Caption = Format(Now, "YYYY-MM-DD HH:NN:SS")

If Format(Now, "NNSS") = "0001" Then
    List1.Clear
End If

If (EasyVil_Cnt < 200) Then
    EasyVil_Cnt = EasyVil_Cnt + 1
    lbl_Cnt.Caption = EasyVil_Cnt
Else
    EasyVil_Cnt = 0
    Call EasyVil_ALIVE
End If
lbl_Cnt.Caption = EasyVil_Cnt

End Sub

Private Sub HomeSock_DataArrival(ByVal bytesTotal As Long)
Dim rMsg As String
Dim B() As Byte
Dim Ret As Integer
Dim Ret1 As Integer
Dim Ret2 As Integer
Dim i, q As Integer
Dim sdata As String
Dim Error_Str As String
ReDim B(bytesTotal - 1)
Dim TArr() As String

'Resp_Falg = False 'µ¥ÀÌÅÍ¸¦ ¼ö½Å¹ÞÀ¸¸é ¼ö½Å´ë±â»óÅÂ¸¦ ÃÊ±âÈ­
HomeSock.GetData B(), vbArray + vbByte, bytesTotal
For i = 0 To bytesTotal - 1
    'Debug.Print b(i)
    rMsg = rMsg & Chr$(B(i))
Next i
'Debug.Print rMsg
Call HomeLogger(" [HomeSock_RCV ] " & rMsg)
    
TArr() = Split(rMsg, "$")
For q = LBound(TArr) To UBound(TArr)
'Debug.Print TArr(q)
Next q
    
Select Case Mid(TArr(2), 5, 2)
    Case 10
        Call HomeLogger(" [HomeSock_RCV ] »óÅÂÃ¼Å© ¿äÃ»")
        sdata = "$version=3.0$" & TArr(3) & "$cmd=11$target=gateway#dongho=" & ezVille_Dong & "&" & ezVille_Ho & "#ip=127.0.0.1#status=0#curtime=" & Format(Now, "YYYYMMDDHHNNSS") & "#hwversion=1.0#swversion=1.1"
        i = LenH(sdata) + LenH("<start=0000&0>")
        sdata = "<start=" & Format(i, "0000") & "&0>" & sdata
        
        'ReDim bData(Len(sData) - 1) As Byte
        'bData = StrConv(sData, vbFromUnicode)
        Call HomeLogger(" [HomeSock_SND ] " & sdata)
        HomeSock.SendData sdata
        EasyVil_Mode = ""
        Call HomeLogger(" [HomeSock_RCV ] »óÅÂÃ¼Å© ÀÀ´ä")
    Case 11
        Call HomeLogger(" [HomeSock_RCV ] Á¶È¸ÀÀ´ä")
    Case 20
        Call HomeLogger(" [HomeSock_RCV ] Á¦¾î¿äÃ»")
    Case 21
        Call HomeLogger(" [HomeSock_RCV ] Á¦¾îÀÀ´ä")
    Case 31 '¼ö½Å
        Call HomeLogger(" [HomeSock_RCV ] ¼ö½Å")
    Case Else
        'Debug.Print TArr(2)
        'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [´ÜÁö¼­¹öÆÐÅ¶ ¼ö½Å] " & rMsg
End Select
    
Exit Sub

Err_P:
    Call HomeLogger(" [Winsock_Kocom_DataArrival Proc] ¿¡·¯³»¿ë : " & Err.Description)

End Sub








