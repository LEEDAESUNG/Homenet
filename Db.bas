Attribute VB_Name = "DBase"
Option Explicit

Public AdoConn_Str As String
Public adoConn As New ADODB.Connection
Public adoHome As New ADODB.Connection

'// DB Open
Public Function DataBaseOpen(ByRef pAdoCon As ADODB.Connection) As Boolean

On Error GoTo Error_Result
    pAdoCon.ConnectionString = AdoConn_Str
    pAdoCon.CursorLocation = adUseClient
    pAdoCon.Open
    DataBaseOpen = True
    Exit Function
Error_Result:

End Function

'// Db Close
Public Sub DataBaseClose(ByRef pAdoConn As ADODB.Connection)

On Error GoTo Error_Result
    pAdoConn.Close
    Set pAdoConn = Nothing
Error_Result:

End Sub

'// DB Open
Public Function HomeDB_Open(ByRef qAdoCon As ADODB.Connection) As Boolean

On Error GoTo Error_Result
    qAdoCon.ConnectionString = AdoHome_Str
    qAdoCon.CursorLocation = adUseClient
    qAdoCon.Open
    HomeDB_Open = True
    FrmGyeyoung.lbl_HomeState.Caption = Format(Now, "YYYY-MM-DD HH:NN:SS") & "    서버 접속 성공"
    Exit Function

Error_Result:
    Call DataLogger(Err.Description)
End Function

'// Db Close
Public Sub HomeDB_Close(ByRef qAdoConn As ADODB.Connection)

On Error GoTo Error_Result
    qAdoConn.Close
    Set qAdoConn = Nothing
Error_Result:

End Sub



Public Sub MakeCSV(lv As ListView, CSVname As String)
    Dim intFileNum As Integer
    Dim ecdata As New ExcelFile
    Dim i, j As Long
    Dim tmpHeader As String
    Dim tmpRS As String

    tmpHeader = ""

    For i = 1 To lv.ColumnHeaders.Count
        If i = 1 Then
            tmpHeader = Trim(lv.ColumnHeaders.Item(1).Text)
        Else
            tmpHeader = tmpHeader & "," & Trim(lv.ColumnHeaders.Item(i).Text)
        End If
    Next i
    
    intFileNum = FreeFile()
    Open CSVname & ".CSV" For Append As #intFileNum
    Print #intFileNum, tmpHeader

    For i = 1 To lv.ListItems.Count
        For j = 1 To lv.ColumnHeaders.Count
            If j = 1 Then
                tmpRS = tmpRS & lv.ListItems(i).Text
            Else
                tmpRS = tmpRS & "," & lv.ListItems(i).SubItems(j - 1)
            End If
        Next j
        'Debug.Print tmpRS
        Print #intFileNum, tmpRS
        tmpRS = ""
    Next i

    Close #intFileNum
    MsgBox "저장이 완료되었습니다."

End Sub


