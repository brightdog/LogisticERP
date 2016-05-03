VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsws_SFMainService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'(fBytes, 0, UBound(fBytes), FLe.Name, "zxb@fjw")
Public Function SOAPUploadFile(ByRef ar_buffer() As Byte, ByVal lng_offset As Long, ByVal lng_length As Long, ByVal str_strFileName As String, ByVal str_strPWD As String) As Boolean

    Dim strContent As String
    Dim objCfgReader As clsReadConfigFile
    Set objCfgReader = New clsReadConfigFile
    strContent = objCfgReader.ReadCfgFile("SoapContent.config", True)
    
    Dim objBase64 As clsvbBase64
    Set objBase64 = New clsvbBase64
    Dim strBase64Conteng As String
    Dim strBase64Content As String
    strBase64Content = objBase64.Base64Encode(ar_buffer)
    
    strContent = Replace(strContent, "|-buffer-|", strBase64Content, 1, 1, vbBinaryCompare)
    strContent = Replace(strContent, "|-strFileName-|", str_strFileName, 1, 1, vbBinaryCompare)
    strContent = Replace(strContent, "|-strPWD-|", str_strPWD, 1, 1, vbBinaryCompare)
    strContent = Replace(strContent, "|-offset-|", lng_offset, 1, 1, vbBinaryCompare)
    strContent = Replace(strContent, "|-length-|", lng_length, 1, 1, vbBinaryCompare)
    
    
    Set objBase64 = Nothing
    Set objCfgReader = Nothing
    
    Dim iSOAPWeb As clsSOAP_Poster
    Set iSOAPWeb = New clsSOAP_Poster
    
    iSOAPWeb.URL = "http://218.17.224.215:8082/SFMainService.svc"
    iSOAPWeb.SOAPAction = "http://tempuri.org/ISFMainService/UploadFile"
    iSOAPWeb.PostData = strContent
    
    Call iSOAPWeb.Send
    
    Debug.Print iSOAPWeb.ReturnData
    
    If iSOAPWeb.ReturnCode = 200 Then
    
        SOAPUploadFile = True
    
    Else
        SOAPUploadFile = False
    End If
    
    
    Set iSOAPWeb = Nothing
    
End Function


'============================== 以下作废！@ 2015-09-06 by wulf ==================================
'=============================== 因为实在是部署得蛋疼 ~~~~ ==================================
''*****************************************************************
''该类是由 Microsoft Office 2003 Web Services Toolkit 创建的。
''
''创建时间: 8/6/2015 08:40:44 AM
''
''说明:
''该类是 Web 服务
'' 的 Visual Basic for Applications 类表示形式，这是由 http://218.17.224.215:8082/SFMainService.svc?wsdl 定义的。
''
''使用:
''将变量声明为新的 clsws_SFMainService，然后编写代码以
''使用该类所提供的方法。
''示例:
'' Dim ExampleVar as New clsws_SFMainService
'' debug.print ExampleVar.wsm_UploadFile(“示例输入”)
''
''有关更多信息，请参阅 Microsoft Office 2003
''Web Services Toolkit 帮助中的“复杂类型”。
''
''对此类中的代码所做的更改可能导致错误的行为。
''
''*****************************************************************
'
''声明专用类变量。
'Private sc_SFMainService As SoapClient30
'Private Const c_WSDL_URL As String = "http://218.17.224.215:8082/SFMainService.svc?wsdl"
'Private Const c_SERVICE As String = "SFMainService"
'Private Const c_PORT As String = "BasicHttpBinding_ISFMainService"
'Private Const c_SERVICE_NAMESPACE As String = "http://tempuri.org/"
'
'Private Sub Class_Initialize()
'    '*****************************************************************
'    '每次实例化该类时都将调用此子例程。
'    '将 sc_ComplexTypes 创建为新的 SoapClient30，然后
'    '用在
'    'http://218.17.224.215:8082/SFMainService.svc?wsdl 中找到的 WSDL 文件实例化 sc_ComplexTypes.mssoapinit2。
'    '*****************************************************************
'
'    Dim str_WSML As String
'    str_WSML = ""
'
'    Set sc_SFMainService = New SoapClient30
'
'    sc_SFMainService.MSSoapInit2 c_WSDL_URL, str_WSML, c_SERVICE, c_PORT, c_SERVICE_NAMESPACE
'    '使用 Internet Explorer 的 LAN 设置中定义的代理服务器，方法是
'    '将 ProxyServer 设置为 <CURRENT_USER>
'    sc_SFMainService.ConnectorProperty("ProxyServer") = "<CURRENT_USER>"
'    '如果 Internet Explorer 被设置为自动检测，则自动检测代理设置，
'    '方法是将 EnableAutoProxy 设置为真
'    sc_SFMainService.ConnectorProperty("EnableAutoProxy") = True
'
'
'End Sub
'
'Private Sub Class_Terminate()
'    '*****************************************************************
'    '将在每次析构该类时调用此子例程。
'    '将 sc_ComplexTypes 设置为 Nothing。
'    '*****************************************************************
'
'    '错误捕获
'    On Error GoTo Class_TerminateTrap
'
'    Set sc_SFMainService = Nothing
'
'Exit Sub
'
'Class_TerminateTrap:
'    SFMainServiceErrorHandler ("Class_Terminate")
'End Sub
'
'Private Sub SFMainServiceErrorHandler(str_Function As String)
'    '*****************************************************************
'    '此子例程是类错误处理程序。当任何类子例程或函数
'    '遇到错误时，都可以从该子例程或函数调用此子例程。然后，它将引发错误并提供
'    '调用子例程或函数的名称。
'    '*****************************************************************
'
'    'SOAP 错误
'    If sc_SFMainService.FaultCode <> "" Then
'        Err.Raise vbObjectError, str_Function, sc_SFMainService.FaultString
'    '非 SOAP 错误
'    Else
'        Err.Raise Err.Number, str_Function, Err.Description
'    End If
'
'End Sub
'
'Public Function wsm_UploadFile(ByVal ar_buffer As Variant, ByVal lng_offset As Long, ByVal lng_length As Long, ByVal str_strFileName As String, ByVal str_strPWD As String) As Boolean
'    '*****************************************************************
'    '从 http://218.17.224.215:8082/SFMainService.svc?wsdl 创建的代理函数。
'    '
'    '“ar_buffer”是一个其元素被定义为 Byte 的数组
'    '有关实现数组的详细信息，
'    '请参阅 Microsoft Office 2003 Web Services Toolkit 帮助中的“复杂类型：数组”。
'    '*****************************************************************
'
'    '错误捕获
'    On Error GoTo wsm_UploadFileTrap
'
'    wsm_UploadFile = sc_SFMainService.UploadFile(ar_buffer, lng_offset, lng_length, str_strFileName, str_strPWD)
'
'Exit Function
'wsm_UploadFileTrap:
'    SFMainServiceErrorHandler "wsm_UploadFile"
'End Function
'
'Public Function wsm_UploadBigFile(ByVal ar_buffer As Variant, ByVal lng_offset As Long, ByVal lng_length As Long, ByVal dcml_totalLength As Double, ByVal str_strFileName As String, ByVal str_strPWD As String) As Boolean
'    '*****************************************************************
'    '从 http://218.17.224.215:8082/SFMainService.svc?wsdl 创建的代理函数。
'    '
'    '“ar_buffer”是一个其元素被定义为 Byte 的数组
'    '有关实现数组的详细信息，
'    '请参阅 Microsoft Office 2003 Web Services Toolkit 帮助中的“复杂类型：数组”。
'    '*****************************************************************
'
'    '错误捕获
'    On Error GoTo wsm_UploadBigFileTrap
'
'    wsm_UploadBigFile = sc_SFMainService.UploadBigFile(ar_buffer, lng_offset, lng_length, dcml_totalLength, str_strFileName, str_strPWD)
'
'Exit Function
'wsm_UploadBigFileTrap:
'    SFMainServiceErrorHandler "wsm_UploadBigFile"
'End Function
'
'Public Function wsm_GetFileList(ByVal str_strPWD As String) As Variant
'    '*****************************************************************
'    '从 http://218.17.224.215:8082/SFMainService.svc?wsdl 创建的代理函数。
'    '
'    '“wsm_GetFileList”是一个其元素被定义为 String 的数组
'    '有关实现数组的详细信息，
'    '请参阅 Microsoft Office 2003 Web Services Toolkit 帮助中的“复杂类型：数组”。
'    '*****************************************************************
'
'    '错误捕获
'    On Error GoTo wsm_GetFileListTrap
'
'    wsm_GetFileList = sc_SFMainService.GetFileList(str_strPWD)
'
'Exit Function
'wsm_GetFileListTrap:
'    SFMainServiceErrorHandler "wsm_GetFileList"
'End Function
'
'Public Function wsm_GetFileLength(ByVal str_strFileName As String, ByVal str_strPWD As String) As Double
'    '*****************************************************************
'    '从 http://218.17.224.215:8082/SFMainService.svc?wsdl 创建的代理函数。
'    '*****************************************************************
'
'    '错误捕获
'    On Error GoTo wsm_GetFileLengthTrap
'
'    wsm_GetFileLength = sc_SFMainService.GetFileLength(str_strFileName, str_strPWD)
'
'Exit Function
'wsm_GetFileLengthTrap:
'    SFMainServiceErrorHandler "wsm_GetFileLength"
'End Function
'
'Public Function wsm_DownloadFile(ByVal str_strFileName As String, ByVal lng_offset As Long, ByVal lng_length As Long, ByVal str_strPWD As String) As Variant
'    '*****************************************************************
'    '从 http://218.17.224.215:8082/SFMainService.svc?wsdl 创建的代理函数。
'    '
'    '“wsm_DownloadFile”是一个其元素被定义为 Byte 的数组
'    '有关实现数组的详细信息，
'    '请参阅 Microsoft Office 2003 Web Services Toolkit 帮助中的“复杂类型：数组”。
'    '*****************************************************************
'
'    '错误捕获
'    On Error GoTo wsm_DownloadFileTrap
'
'    wsm_DownloadFile = sc_SFMainService.DownLoadFile(str_strFileName, lng_offset, lng_length, str_strPWD)
'
'Exit Function
'wsm_DownloadFileTrap:
'    SFMainServiceErrorHandler "wsm_DownloadFile"
'End Function
'
'Public Function wsm_delFile(ByVal str_strFileName As String, ByVal str_strPWD As String) As Boolean
'    '*****************************************************************
'    '从 http://218.17.224.215:8082/SFMainService.svc?wsdl 创建的代理函数。
'    '*****************************************************************
'
'    '错误捕获
'    On Error GoTo wsm_delFileTrap
'
'    wsm_delFile = sc_SFMainService.delFile(str_strFileName, str_strPWD)
'
'Exit Function
'wsm_delFileTrap:
'    SFMainServiceErrorHandler "wsm_delFile"
'End Function
'
'Public Function wsm_IsUploadFile(ByVal str_strPWD As String) As Boolean
'    '*****************************************************************
'    '从 http://218.17.224.215:8082/SFMainService.svc?wsdl 创建的代理函数。
'    '*****************************************************************
'
'    '错误捕获
'    On Error GoTo wsm_IsUploadFileTrap
'
'    wsm_IsUploadFile = sc_SFMainService.IsUploadFile(str_strPWD)
'
'Exit Function
'wsm_IsUploadFileTrap:
'    SFMainServiceErrorHandler "wsm_IsUploadFile"
'End Function
'
'Public Function wsm_GetUploadTime(ByVal str_strPWD As String) As Long
'    '*****************************************************************
'    '从 http://218.17.224.215:8082/SFMainService.svc?wsdl 创建的代理函数。
'    '*****************************************************************
'
'    '错误捕获
'    On Error GoTo wsm_GetUploadTimeTrap
'
'    wsm_GetUploadTime = sc_SFMainService.GetUploadTime(str_strPWD)
'
'Exit Function
'wsm_GetUploadTimeTrap:
'    SFMainServiceErrorHandler "wsm_GetUploadTime"
'End Function
'
'
'
