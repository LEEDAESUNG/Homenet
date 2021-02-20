VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  '없음
   ClientHeight    =   6000
   ClientLeft      =   9990
   ClientTop       =   5310
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   6000
   ScaleWidth      =   9000
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1410
      TabIndex        =   8
      Top             =   3780
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   120
      Top             =   5400
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "본 프로그램은 프로그램등록법에 의거 정식등록 되어있으므로, 무단복제하는 경우에는 저작권의 침해가 되므로 주의 바랍니다. "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   2670
      TabIndex        =   7
      Top             =   4785
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "JWT ⓒ 2003-2015 All Right Reserved. See the patent and legal notice in the about box."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   780
      TabIndex        =   6
      Top             =   5310
      Width           =   7860
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "Warning"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9615
      TabIndex        =   5
      Top             =   1995
      Width           =   1455
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   8775
   End
   Begin VB.Label lblLicenseTo 
      BackStyle       =   0  '투명
      Caption         =   "LicenseTo"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  '투명
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   645
      TabIndex        =   2
      Top             =   1335
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "HomeNet Manager™  for LPR System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   7095
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  '투명
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   645
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim tm_cnt As Integer

Private Sub Form_Load()

If App.PrevInstance = True Then
    End
End If

Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2

lblVersion.Caption = "Version " & App.Major & " . " & App.Minor & " . " & App.Revision
lblCompany.Caption = App.CompanyName
lblLicenseTo.Caption = App.FileDescription
IniFileName$ = App.Path & "\HomeNet.ini"
Doc_Path_Name$ = App.Path & "\Doc\"
Report_Path_Name$ = App.Path & "\Data\"

'전역변수 초기화
Call CFG_Init

ProgressBar1.Max = 100
Timer1.Enabled = True

End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()

On Error Resume Next

    If (tm_cnt <= 99) Then
        tm_cnt = tm_cnt + 1
        ProgressBar1.value = tm_cnt
    Else
        Timer1.Enabled = False
        Unload Me
        
        Select Case HomeNetMode
            Case 1
                FrmHyundae.Show 0
            Case 2
                FrmGyeyoung.Show 0
            Case 3
                FrmEZVille.Show 0
            Case 4
                FrmKocom.Show 0
            Case 5
                FrmCommax.Show 0
            Case 6
                FrmIControls.Show 0
            Case 7
                FrmKDOne.Show 0
            Case 8
                FrmLGElectron.Show 0
            Case 9
                FrmGyeyoungTCP.Show 0
            Case 10
                FrmHyundae_Linux.Show 0
            Case 11
                FrmMaxuracy.Show 0
            Case 12
                FrmHomeclever.Show 0
            Case 13
                frmGyeyoungRS232.Show 0
                
                
        End Select
    End If

End Sub

