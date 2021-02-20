Attribute VB_Name = "LPR"
Option Explicit
Public Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const MAX_PATH = 260
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000


Public Function Get_Process() As Boolean
Dim cb As Long
Dim cbNeeded As Long
Dim NumElements As Long
Dim ProcessIDs() As Long
Dim cbNeeded2 As Long
Dim NumElements2 As Long
Dim Modules(1 To 200) As Long
Dim lRet As Long
Dim ModuleName As String
Dim nSize As Long
Dim hProcess As Long
Dim i As Long
Dim tmp As String

cb = 8
cbNeeded = 96
Do While cb <= cbNeeded
    cb = cb * 2
    ReDim ProcessIDs(cb / 4) As Long
    lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
Loop

NumElements = cbNeeded / 4
For i = 1 To NumElements
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
    If hProcess <> 0 Then
        lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
        If lRet <> 0 Then
            ModuleName = Space(MAX_PATH)
            nSize = 500
            lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
            tmp = Left(ModuleName, lRet)
            If (Right(tmp, 11) = "Navicat.exe") Then
                Get_Process = True
                Exit Function
            End If
        End If
    End If
    lRet = CloseHandle(hProcess)
Next
Get_Process = False

End Function

Public Sub Data_ReSearch(size As Integer)
End Sub

Public Sub Data_ReSearch_Unload()
End Sub

Public Function Slash_Conv(str As String) As String
Dim i As Integer
Dim tmp As String
Dim Ret As Boolean

tmp = "\\\\"

For i = 3 To LenH(str) Step 1
    If (Mid(str, i, 1) = "\") Then
        tmp = tmp & "\\" & Mid(str, i, 1)
    Else
        tmp = tmp & Mid(str, i, 1)
    End If
Next i

Slash_Conv = tmp

End Function

