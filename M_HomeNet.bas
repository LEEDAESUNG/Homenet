Attribute VB_Name = "M_HomeNet"
Public Sub HomeNet_Proc(sData As String)
Dim i As Integer
    
    
On Error GoTo Err_Proc
    
    HomeNet_Dong = Left(sData, 4)
    HomeNet_Ho = Mid(sData, 5, 4)
    i = Len(sData)
    HomeNet_CarNo = Mid(sData, 9, (i - 8))

Exit Sub

Err_Proc:
    Call HomeLogger("[HomeNet Proc] " & Err.Description)

End Sub

