VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmIControls 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "FrmIControls"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   19170
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
      Left            =   13620
      TabIndex        =   10
      Top             =   540
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   990
      Top             =   6300
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
      Left            =   17820
      TabIndex        =   5
      Top             =   6390
      Width           =   1065
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
      ItemData        =   "FrmIControls.frx":0000
      Left            =   60
      List            =   "FrmIControls.frx":0002
      TabIndex        =   4
      Top             =   1650
      Width           =   19050
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
      Left            =   17850
      TabIndex        =   3
      Top             =   90
      Width           =   1095
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
      TabIndex        =   2
      Text            =   "12121"
      Top             =   990
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
      TabIndex        =   1
      Text            =   "12121"
      Top             =   480
      Width           =   1755
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
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1755
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
      Left            =   14550
      TabIndex        =   9
      Top             =   180
      Width           =   3075
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   60
      X2              =   19110
      Y1              =   1530
      Y2              =   1530
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
      TabIndex        =   8
      Top             =   1020
      Width           =   1515
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
      TabIndex        =   7
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
      Left            =   180
      TabIndex        =   6
      Top             =   150
      Width           =   1305
   End
End
Attribute VB_Name = "FrmIControls"
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
        Call IControls_Proc(Trim(txt_Dong.Text), Trim(txt_Ho.Text), "01°¡1234")
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

Call HomeLogger("[HomeNet Program ] IControls HomeNet Start..!!")

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
    
    '¾ÆÀÌÄÜÆ®·Ñ½º ¼¼´ëÅëº¸
    Call IControls_Proc(HomeNet_Dong, HomeNet_Ho, HomeNet_CarNo)

Exit Sub

Err_P:
    Call HomeLogger(" [HostSock UDP DataArrival]  " & Err.Description)

End Sub

Private Sub HomeSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call HomeLogger("[HomeSock UDP Error]  " & Description)
End Sub

Private Sub Timer1_Timer()
lbl_Date.Caption = Format(Now, "YYYY-MM-DD HH:NN:SS")

If Format(Now, "NNSS") = "0001" Then
    List1.Clear
End If

End Sub

Public Sub IControls_Proc(Dong As String, Ho As String, CarNo As String)
Dim Socket_Data_len As String * 8
Dim year As Integer
Dim mon As Integer
Dim day As Integer
Dim hour As Integer
Dim min As Integer
Dim sec As Integer
Dim DQ As String


On Error GoTo Err_Proc

DQ = """"

Socket_Data = "<?xml version=" & DQ & "1.0" & DQ & " encoding=" & DQ & "utf-8" & DQ & "><imap ver=" & DQ & "1.0" & DQ & " address=" & DQ & Local_IP & DQ & " sender=" & DQ & "ÁÖÂ÷°üÁ¦ ÄÁÆ®·Ñ·¯" & DQ & "><service type=" & DQ & "request" & DQ & " name=" & DQ & "car_move_info" & DQ & "><destination name=" & DQ & "homedev" & DQ & " id_high=" & DQ & Val(Dong) & DQ & " id_low=" & DQ & Val(Ho) & DQ & "/><move_info>" & DQ & "in" & DQ & "</move_info><params car_num=" & DQ & Trim(CarNo) & DQ & " loc_num=" & DQ & "ÀÔ±¸" & DQ & " message=" & DQ & "null" & DQ & "/></service></imap>"
'Socket_Data_len = Len(Socket_Data)
'Socket_Data = Socket_Data_len & Socket_Data
Call HomeLogger("[IControls_Proc]  SND : " & Socket_Data)
Call Socket_Connect

Exit Sub

Err_Proc:
    Call HomeLogger("[IControls_Proc]  " & Err.Description)
End Sub

Public Sub Socket_Connect()
Dim bdata() As Byte

On Error GoTo Err_P
    
    Call HomeLogger("[HomeAlarm_SocketConnect] ¿¬°á½Ãµµ : " & HomeNet_IP & " " & HomeNet_Port)
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
    
    sData = Socket_Data
    ReDim bdata(Len(sData) - 1) As Byte
    bdata = StrConv(sData, vbFromUnicode)
    HomeSock.SendData bdata
    Socket_Data = ""

Exit Sub

Err_P:
Call HomeLogger("[HomeSock_Connect ÇÁ·Î½ÃÁ®] ¿¡·¯³»¿ë : " & Err.Description)
End Sub

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

Call HomeLogger("[IControls_Proc]  RCV : " & rMsg)

HomeSock.Close

Exit Sub

Err_P:
    Call HomeLogger("[IControls_Proc RCV] ¿¡·¯³»¿ë : " & Err.Description)
End Sub







