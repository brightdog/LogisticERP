VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKillProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0
Private Const SW_RESTORE = 9
Private Const SW_SHOW = 5
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPheaplist = &H1
Private Const TH32CS_SNAPthread = &H4
Private Const TH32CS_SNAPmodule = &H8
Private Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule
Private Const MAX_PATH As Integer = 260
Private Const PROCESS_TERMINATE = &H1
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

Public Sub KillProcess(ByVal ProcessName As String)
Dim i As Long
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim exename As String
Dim hand As Long, theloop As Long

snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0) '：获得进程“快照”的句柄
proc.dwSize = Len(proc)
theloop = ProcessFirst(snap, proc) '：获取第一个进程，并得到其返回值
i = 0
While theloop <> 0 '：当返回值非零时继续获取下一个进程
exename = proc.szExeFile
Debug.Print Split(exename, ".exe", 2, vbBinaryCompare)(0)

If Split(LCase(exename), ".exe", 2, vbBinaryCompare)(0) = ProcessName Then
hand = OpenProcess(PROCESS_TERMINATE, True, proc.th32ProcessID) '：获取进程句柄
TerminateProcess hand, 0  '：关闭进程
End If
theloop = ProcessNext(snap, proc)
Wend
CloseHandle snap '：关闭进程“快照”句柄
End
End Sub
