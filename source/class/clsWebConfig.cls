VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWebConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const APP_CONFIG_PATH As String = "\config\"

Public Function ReadProperty(ByVal NodeName As String) As String
    Dim doc     As New MSXML2.DOMDocument60
    Dim Node    As IXMLDOMNode

    On Error Resume Next

    doc.Load App.path & APP_CONFIG_PATH & "web.config.xml"

    Set Node = doc.selectSingleNode("root/config/" & NodeName)

    If Not Node Is Nothing Then
        ReadProperty = Node.Text
    End If

End Function

'Public Function LoadXml() As String()
'
'    Dim strResult() As String
'
'    Dim doc As DOMDocument
'    Set doc = New DOMDocument
'    Dim Nodes    As IXMLDOMNodeList
'
'    doc.Load APP_CONFIG_PATH & "web.config.xml"
'
'    Set Nodes = doc.selectNodes("root")
'
'    If Nodes.Length > 1 Then
'        If Not Nodes Is Nothing Then
'            'ReadProperty = getAttribute(Node, AttrName)
'
'
'
'
'        Else
'            ReDim strResult(0, 0)
'        End If
'
'    Else
'        ReDim strResult(0, 0)
'    End If
'
'    LoadXml = strResult
'End Function

'没时间写了，XML文件就手工写算了。
'Public Function SaveXml(ByRef arr() As String) As String()
'    Dim doc As DOMDocument
'    Set doc = New DOMDocument
'    Dim Nodes    As IXMLDOMNodeList
'
'    doc.Load APP_CONFIG_PATH & "web.config.xml"
'
'    Set Nodes = doc.selectNodes("root")
'
'       If Not Nodes Is Nothing Then
'        Dim root As MSXML2.IXMLDOMNode
'        Set root = doc.createNode(1, "root", "")
'        doc.appendChild root
'        doc.save (APP_CONFIG_PATH & "web.config.xml")
'       End If
'
'    Dim i, j As Integer
'
'    For i = 0 To UBound(arr, 1)
'
'
'        For j = 0 To UBound(arr, 2)
'
'
'        If arr(i, j) Then
'
'
'
'        Else
'
'
'        End If
'
'
'    Next
'
'End Function
