VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCONN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const CONNECTION_FILE As String = "conn.cfg"
Public Function OpenConn() As ADODB.Connection
    Dim Conn As ADODB.Connection
    Set Conn = New ADODB.Connection
    Dim strConnectionString As String
    
    strConnectionString = GetConnection()
    
    Conn.ConnectionString = strConnectionString
    Call Conn.Open

    Set OpenConn = Conn
End Function


Private Function GetConnection() As String
        '<EhHeader>
        On Error GoTo GetConnection_Err
        '</EhHeader>

        Dim Fso As Scripting.FileSystemObject
100     Set Fso = New Scripting.FileSystemObject
        Dim strTmp As String
102     strTmp = Fso.OpenTextFile(App.path & "\" & CONNECTION_FILE, ForReading, False, TristateFalse).ReadAll()
104     If Trim(strTmp) <> "" Then
        Dim objEncrypt As clsEncrypt
106     Set objEncrypt = New clsEncrypt
            
            GetConnection = objEncrypt.Decode(strTmp)
            
        Else
    
108         WriteLog "连接字符串为空！"
    
        End If
    
110     Set Fso = Nothing

        '<EhFooter>
        Exit Function

GetConnection_Err:
        WriteLog "连接配制文件不存在！"
        '</EhFooter>
End Function




