Attribute VB_Name = "WMCopy"
Option Explicit

Public Const WM_LANE1_HANDLE = "01"
Public Const WM_LANE2_HANDLE = "02"
Public Const WM_LANE3_HANDLE = "03"
Public Const WM_LANE4_HANDLE = "04"

Public Const WM_LANE1_CARNUM = "11"
Public Const WM_LANE2_CARNUM = "12"
Public Const WM_LANE3_CARNUM = "13"
Public Const WM_LANE4_CARNUM = "14"

Public Const WM_LANE1_WATCHDOG_ACK = "21"
Public Const WM_LANE2_WATCHDOG_ACK = "22"
Public Const WM_LANE3_WATCHDOG_ACK = "23"
Public Const WM_LANE4_WATCHDOG_ACK = "24"

Public Const WM_LANE1_LOADING = "31"
Public Const WM_LANE2_LOADING = "32"
Public Const WM_LANE3_LOADING = "33"
Public Const WM_LANE4_LOADING = "34"

Public Const WM_LANE1_CAMERA_ERR = "41"
Public Const WM_LANE2_CAMERA_ERR = "42"
Public Const WM_LANE3_CAMERA_ERR = "43"
Public Const WM_LANE4_CAMERA_ERR = "44"


Public Const WM_HOST_HANDLE = "51"
Public Const WM_FEE1_HANDLE = "52"
Public Const WM_FEE2_HANDLE = "53"

Public Const WM_WATCHDOG_POLL = "99"

Public LANE1_Handle As Long
Public LANE2_Handle As Long
Public LANE3_Handle As Long
Public LANE4_Handle As Long

Global gHW As Long


Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMsg Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


 Public Type COPYDATASTRUCT
              dwData As Long
              cbData As Long
              lpData As Long
 End Type

Public Const GWL_WNDPROC = (-4)
Public Const WM_COPYDATA = &H4A
Global lpPrevWndProc As Long


Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As _
         Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub Hook()
   lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
    Dim Temp As Long
    Temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

'Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
'    If uMsg = WM_COPYDATA Then
'        Call mySub(lParam)
'    End If
'    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
'
'End Function
'
'
'
'Public Function SendMess(ByVal Mess As String, TrHwnd As Long)
'    If TrHwnd = 0 Then Exit Function
'
'    Dim cds As COPYDATASTRUCT
'    Dim ThWnd As Long, Sownd As Long
'    Dim buf(1 To 255) As Byte
'    Dim i As Long
'
'    Dim strsz As Integer
'    Sownd = TrHwnd
'    ThWnd = TrHwnd
'    strsz = LenB(StrConv(Mess, vbFromUnicode)) '' 한글 2바이트 영문은 1바이트
'    Call CopyMemory(buf(1), ByVal Mess, strsz)
'    cds.dwData = 3
'    cds.cbData = strsz + 1
'    cds.lpData = VarPtr(buf(1))
'    'i = SendMessage(ThWnd, WM_COPYDATA, Sownd, cds)
'    i = SendMsg(ThWnd, WM_COPYDATA, Sownd, cds)
'End Function
'
'
