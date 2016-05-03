Attribute VB_Name = "modProc"
'Write By 老吴

Option Explicit
Private Declare Function GetComputerName _
                Lib "kernel32" _
                Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Dim StopService As Boolean

'=============================
Private Declare Function CreateToolhelpSnapshot _
                Lib "kernel32" _
                Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst _
                Lib "kernel32" _
                Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext _
                Lib "kernel32" _
                Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess _
                Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess _
                Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle _
                Lib "kernel32" (ByVal hObject As Long) As Long
Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Const TH32CS_SNAPheaplist = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPthread = &H4
Const TH32CS_SNAPmodule = &H8
Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule
'=============================

Dim mstrComputerName As String

Public Function GetProcessCountbyName(ByVal strName As String) As Integer

    Dim dic As Scripting.Dictionary
    Set dic = GetAllProcessName
    
    
    
    If dic.Exists(strName) Then
    
        GetProcessCountbyName = dic.Item(strName)
    
    Else
    
        GetProcessCountbyName = -1
    
    End If

End Function


Public Function GetAllProcessName() As Scripting.Dictionary
    Dim i As Long, lPid As Long
    Dim Proc As PROCESSENTRY32
    Dim hSnapShot As Long
    Dim strResult As String
    strResult = ""
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPall, 0) '获得进程“快照”的句柄
    Proc.dwSize = Len(Proc)
    lPid = ProcessFirst(hSnapShot, Proc) '获取第一个进程的PROCESSENTRY32结构信息数据
    i = 0
    
    Dim dicProcess As Scripting.Dictionary
    Set dicProcess = New Scripting.Dictionary
    
    Do While lPid <> 0 '当返回值非零时继续获取下一个进程
        'Debug.Print Left(Proc.szExeFile, InStr(1, Proc.szExeFile, Chr(0)) - 1)
        'strResult = strResult & Left(Proc.szExeFile, InStr(1, Proc.szExeFile, Chr(0)) - 1) & "|"
        'ListView1.ListItems.Add , "a" & i, Hex(Proc.th32ProcessID) '将进程ID添加到ListView1第一列
        'ListView1.ListItems("a" & i).SubItems(1) = Proc.szExeFile '将进程名添加到ListView1第二列
        
        Dim tmpName As String
        
        tmpName = Left(Proc.szExeFile, InStr(1, Proc.szExeFile, Chr(0)) - 1)
        
        If Not dicProcess.Exists(tmpName) Then
        
            dicProcess.Add tmpName, 1
        Else
        
            dicProcess.Item(tmpName) = dicProcess.Item(tmpName) + 1
        
        End If

        lPid = ProcessNext(hSnapShot, Proc) '循环获取下一个进程的PROCESSENTRY32结构信息数据
    Loop

    CloseHandle hSnapShot '关闭进程“快照”句柄
    
    
    Set GetAllProcessName = dicProcess
    
'    Dim v As Variant
'
'    For Each v In dicProcess.Keys
'
'        strResult = strResult & v & "(" & dicProcess.Item(v) & ")|"
'
'    Next
'
'    Set dicProcess = Nothing
'    GetProcessName = strResult
End Function

