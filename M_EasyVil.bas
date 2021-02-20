Attribute VB_Name = "M_EasyVil"
Option Explicit

Public EasyVil_Mode As String
Public EasyVil_Cnt As Integer


Public Sub EasyVil_ALIVE()
Dim i As Integer
Dim sData As String

On Error GoTo Err_P

    EasyVil_Mode = "ALIVE"
    If FrmEZVille.HomeSock.State = 7 Then
        sData = "$version=3.0$copy=00-0000$cmd=10$dongho=" & ezVille_Dong & "&" & ezVille_Ho & "$target=server"
        i = LenH(sData) + LenH("<start=0000&0>")
        sData = "<start=" & Format(i, "0000") & "&0>" & sData
        
        Call HomeLogger(" [HomeSock_SND ] " & sData)
        Call HomeLogger(" [HomeSock_SND ] HomeNet 상태체크 요청")
        FrmEZVille.HomeSock.SendData sData
    Else
        EasyVil_Mode = ""
        Call HomeLogger(" [EasyVil_Alive] Alive Check SND Failure..!!")
        Call Socket_Connect_EasyVil
        Call HomeLogger(" [EasyVil_ReConnect] Reconnect..!!")
    End If

Exit Sub

Err_P:
     Call HomeLogger("[EasyVil_ALIVE Proc] :" & Err.Description)

End Sub

Public Sub EasyVil_Alarm(inout As Integer, CarNo As String, tmpDong As String, tmpHo As String)
Dim i As Integer
Dim sData As String
Dim Car() As Byte
Dim sDong, sHo As String
Dim tmpstr As String * 12

    EasyVil_Mode = "ALARM"
    If FrmEZVille.HomeSock.State = 7 Then
        sData = "$version=3.0$cmd=30$dongho=" & ezVille_Dong & "&" & ezVille_Ho & "$target=parking#param=#dongho=" & Val(tmpDong) & "&" & Val(tmpHo) & "#inout=0#carno=" & Right(Trim(CarNo), 4) & "#time=" & Format(Now, "YYYYMMDDHHNNSS")
        i = LenH(sData) + LenH("<start=0000&0>")
        sData = "<start=" & Format(i, "0000") & "&0>" & sData
        Call HomeLogger(" [HomeSock_SND ] " & sData)
        FrmEZVille.HomeSock.SendData sData
    Else
        EasyVil_Mode = ""
        Call HomeLogger(" [EasyVil_Alarm Proc]     Alarm Fail : Discoonect.!!")
    End If

End Sub


Public Sub Socket_Connect_EasyVil()
    Dim bdata() As Byte
    
On Error GoTo Err_P
    
With FrmEZVille
    '.HomeSock.Close
    If (.HomeSock.State <> sckClosed) Then
        .HomeSock.Close
    End If
    .HomeSock.Connect HomeNet_IP, HomeNet_Port
    Call HomeLogger(" [ezVille_Connection]" & " HomeNet_IP : " & HomeNet_IP & " HomeNet_Port : " & HomeNet_Port)
End With

Exit Sub

Err_P:
    Call HomeLogger(" [ezVille_Connect Proc] 에러내용 : " & Err.Description)

End Sub

