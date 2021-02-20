VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmKocom 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "ÄÚÄÞ ¼¼´ëÅëº¸"
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
      ItemData        =   "FrmKocom.frx":0000
      Left            =   30
      List            =   "FrmKocom.frx":0002
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
   Begin VB.Label lbl_KocomCnt 
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
      Caption         =   "Kocom CNT : "
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
      Width           =   1275
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
Attribute VB_Name = "FrmKocom"
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
    Call HomeLogger("HomeAlarm Test Send : " & "12°¡1234" & "  " & Val(txt_Dong) & "  " & Val(txt_Ho))
    Call Kocom_Alarm(0, "12°¡1234", Val(txt_Dong), Val(txt_Ho))
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

Call HomeLogger("[HomeNet Program] KOCOM HomeNet Start..!!")

'ÄÚÄÞ ¹ÙÀÎµù==============================================================================
Kocom_Cnt = 0
With KHeader
    .KEY(0) = &H78
    .KEY(1) = &H56
    .KEY(2) = &H34
    .KEY(3) = &H12
    .MSGTYPE(0) = &H0
    .MSGTYPE(1) = &H0
    .MSGTYPE(2) = &H0
    .MSGTYPE(3) = &H11
    For i = 0 To 3
        .MSGLENGTH(i) = &H0
        .TOWN(i) = &H0
        .Dong(i) = &H0
        .Ho(i) = &H0
        .Reserved(i) = &H0
    Next
End With
Call Kocom_BIND(HomeNet_ID, HomeNet_PW)
'ÄÚÄÞ ¹ÙÀÎµù==============================================================================

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
    
    'ÄÚÄÞ ¼¼´ëÅëº¸
    'Kocom_Alarm(inout As Integer, CarNo As String, tmpDong As Integer, tmpHo As Integer)
    Call Kocom_Alarm(0, Trim(HomeNet_CarNo), Val(HomeNet_Dong), Val(HomeNet_Ho))

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

If (Kocom_Cnt < 50) Then
    Kocom_Cnt = Kocom_Cnt + 1
Else
    If Kocom_Mode = "" Or Kocom_Mode = "BIND" Then
        Kocom_Cnt = 0
        Call Kocom_ALIVE
    End If
End If
lbl_KocomCnt.Caption = Kocom_Cnt

End Sub

Private Sub HomeSock_Connect()
Dim sData As String
Dim bdata() As Byte
Dim i As Integer

On Error GoTo Err_P

Select Case Kocom_Mode
    Case "BIND"
        With KHeader     'Çì´õÀü¼Û
            HomeSock.SendData .KEY()
            HomeSock.SendData .MSGTYPE()
            HomeSock.SendData .MSGLENGTH()
            HomeSock.SendData .TOWN()
            HomeSock.SendData .Dong()
            HomeSock.SendData .Ho()
            HomeSock.SendData .Reserved()
        End With
        With KBind
            HomeSock.SendData .HomeVersion()
            HomeSock.SendData .nKind()
            HomeSock.SendData .nVersion()
            HomeSock.SendData .SzId()
            HomeSock.SendData .SzPass()
        End With
End Select

Exit Sub

Err_P:
Call HomeLogger(Format(Now, "yyyy-mm-dd hh:nn:ss") & " [HomeSock_Connect ÇÁ·Î½ÃÁ®] ¿¡·¯³»¿ë : " & Err.Description)
End Sub

Private Sub HomeSock_DataArrival(ByVal bytesTotal As Long)
    Dim rMsg As String
    Dim B() As Byte
    Dim Ret As Integer
    Dim i As Integer
    Dim sData As String

On Error GoTo Err_P
    
    ReDim B(bytesTotal - 1)
    'Debug.Print bytesTotal
    HomeSock.GetData B(), vbArray + vbByte, bytesTotal
'    For i = 0 To bytesTotal - 1
'        Debug.Print ("&H" & Hex(B(i)))
'    Next i
'    Debug.Print "-------------"
    
    For i = 0 To bytesTotal - 1
        If (B(i) >= &H80) Then
            rMsg = rMsg & Chr$(Val("&H" & Hex(B(i)) & Hex(B(i + 1))))
            i = i + 1
        Else
            rMsg = rMsg & Chr$(B(i))
        End If
    Next i
    Call HomeLogger("[Kocom_Proc]  RCV Data" & rMsg)
    
    If "&H" & Hex(B(3)) & Hex(B(2)) & Hex(B(1)) & Hex(B(0)) = &H12345678 Then
        'ÄÁ³Ø¼ÇÀÌ ²÷¾îÁø »óÅÂ
        If "&H" & Hex(B(31)) & Hex(B(30)) = &H2710 Then
            
        End If
        
        Select Case "&H" & Hex(B(31)) & Hex(B(30)) & Hex(B(29)) & Hex(B(28))
            Case &H0
                    'SUCCESS
                    Call HomeLogger("[Winsock_Kocom_DataArrival Proc] " & Kocom_Mode & " Success..!!")
            Case &H1
                    'AUTH_FAIL_ID
                    Call HomeLogger("[Winsock_Kocom_DataArrival Proc] " & Kocom_Mode & " AUTH_FAIL_ID..!!")
            Case &H2
                    'AUTH_FAIL_PASS
                    Call HomeLogger("[Winsock_Kocom_DataArrival Proc] " & Kocom_Mode & " AUTH_FAIL_PASS..!!")
            Case &H3
                    'OFF_HOME
                    Call HomeLogger("[Winsock_Kocom_DataArrival Proc] " & Kocom_Mode & " OFF_HOME..!!")
            Case &H7
                    'RE_TRY_DATA
                    Call HomeLogger("[Winsock_Kocom_DataArrival Proc] " & Kocom_Mode & " RE_TRY_DATA..!!")
            Case &H1100
                    'OVERFLOW_Q_SIZE
                    Call HomeLogger("[Winsock_Kocom_DataArrival Proc] " & Kocom_Mode & " OVERFLOW_Q_SIZE..!!")
            Case Else
                    Call HomeLogger("[Winsock_Kocom_DataArrival Proc] " & Kocom_Mode & " OK..!!")
        End Select
    End If
    Kocom_Mode = ""
    
Exit Sub

Err_P:
    Call HomeLogger("[Winsock_Kocom_DataArrival Proc] ¿¡·¯³»¿ë : " & Err.Description)

End Sub

Private Sub HomeSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call HomeLogger("[HomeSock TCP Error]  " & Description)
End Sub






