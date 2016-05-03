VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ADSLConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mstrNetWorkState As String '纪录当前ADSL的拨号环境，每次拨号前都检测一下。
Private Const CONNECTSTATEFILE As String = "\ConnectType.CFG"

Private Sub InitNetWorkState()

    Dim FSo As Scripting.FileSystemObject
    Set FSo = New Scripting.FileSystemObject
    
    If FSo.FileExists(App.path & CONNECTSTATEFILE) Then
    
        mstrNetWorkState = FSo.OpenTextFile(App.path & CONNECTSTATEFILE, ForReading, False, TristateFalse).ReadLine
        '文件可能为空，那这个变量就为空字符串
    Else
    
        mstrNetWorkState = ""
    
    End If
    Set FSo = Nothing
    
End Sub

Public Function Connect() As Boolean
    Dim obj As Object

    Select Case mstrNetWorkState

        Case "ROUTER"
            Set obj = New clsADSLByRouter
            
        Case "MODEM"
    
            Set obj = New clsADSLByModem
        Case Else
            Set obj = New clsADSLByRouter
            
    End Select

    Call obj.Connect
            
    Set obj = Nothing
End Function

Public Function Disconnect() As Boolean
    Dim obj As Object

    Select Case mstrNetWorkState

        Case "ROUTER"
            Set obj = New clsADSLByRouter
            
        Case "MODEM"
    
            Set obj = New clsADSLByModem

        Case Else
            Set obj = New clsADSLByRouter
    End Select

    Call obj.Disconnect
            
    Set obj = Nothing

End Function


Private Sub Class_Initialize()

    Call InitNetWorkState


End Sub



