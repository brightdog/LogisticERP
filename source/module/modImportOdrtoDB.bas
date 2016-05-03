Attribute VB_Name = "modImportOdrtoDB"
Option Explicit
Public Const gstrOrderFieldHeader As String = "包装单号码|出货工作单号码|地址|姓名|收货人电话|件数|毛重|货品总金额|备注|签收日期" '签收人||承运商号码|运单号"

Public Const gstrOrderTableFieldHeader As String = "ThdPkgNO|ThdOrderOutWorkNO|ReceiverAddress|ReceiverName|ReceiverMobile|PkgNum|PkgWeight|InsurePrice|Remark|ReceiveDateTime"

Public Const gstrExpressFieldHeader As String = "承运商号码|运单号码"

Public Function ImportOrderDataToDB(ByVal strRawData As String) As String
    Dim i As Integer
    Dim arrLine()  As String
    
    Dim strResult As String
    
    arrLine = Split(strRawData, vbCrLf, -1, vbBinaryCompare)
    
    If UBound(arrLine) > 0 Then '至少要2行，一行列头，一行正文
    
        Dim iLine As Integer
        
        Dim arrOrderFieldHeader() As String
        Dim arrOrderTableFieldHeader() As String
        Dim arrExpressFieldHeader() As String
        
        arrOrderFieldHeader = Split(gstrOrderFieldHeader, "|", -1, vbBinaryCompare)
        
        arrOrderTableFieldHeader = Split(gstrOrderTableFieldHeader, "|", -1, vbBinaryCompare)
        
        arrExpressFieldHeader = Split(gstrExpressFieldHeader, "|", -1, vbBinaryCompare)
        
        Dim strData As String
        strData = ""
        
        Dim dicFieldValueOrder As Scripting.Dictionary
        Set dicFieldValueOrder = New Scripting.Dictionary
                        
        Dim strFieldOrder As String
        Dim dicExpressFieldValue As Scripting.Dictionary
        Set dicExpressFieldValue = New Scripting.Dictionary

        For i = 0 To UBound(arrExpressFieldHeader)

            dicExpressFieldValue.Add arrExpressFieldHeader(i), ""
            
        Next

        For iLine = 0 To UBound(arrLine)

            If arrLine(iLine) <> "" Then
                Dim v As Variant
            
                Dim arrField() As String
            
                arrField = Split(arrLine(iLine), vbTab, -1, vbBinaryCompare)
            
                Dim dicFieldValue As Scripting.Dictionary
                Set dicFieldValue = New Scripting.Dictionary
                
                For i = 0 To UBound(arrOrderTableFieldHeader)

                    dicFieldValue.Add arrOrderTableFieldHeader(i), ""
            
                Next

                '为每一行数据初始化字典对象，检查每一个值是否都有，否则就在导入之前报错提示。
                '而且也就是第一行的时候，需要确定字段顺序，后续只要检查是否有值应该就可以了。
                If iLine = 0 Then
                    '确认字段顺序，以及字段是否有存在
                
                    For i = 0 To UBound(arrField)
                        Dim strMappingField As String
                        strMappingField = GetMappingField(arrField(i), arrOrderFieldHeader, arrOrderTableFieldHeader)

                        If strMappingField <> "" Then
                            If dicFieldValue.Exists(strMappingField) Then
                                dicFieldValueOrder.Item(strMappingField) = i '存放每一个字段的顺序
                            Else
                                'WriteLog "有多余的字段:" & arrField(i)
                                '先扔一下不处理，大不了不导入就是了。继续检查其他的。
                            End If
                        End If
                            
                        If dicExpressFieldValue.Exists(arrField(i)) Then
                            dicExpressFieldValue.Item(arrField(i)) = i '存放每一个字段的顺序
                        Else
                            'WriteLog "有多余的字段:" & arrField(i)
                            '先扔一下不处理，大不了不导入就是了。继续检查其他的。
                        End If

                    Next

                    '                For i = 0 To UBound(arrExpressFieldHeader)
                    '
                    '                    dicExpressFieldValue.Add arrExpressFieldHeader(i), ""
                    '
                    '                Next
                    For Each v In dicFieldValueOrder.keys
                
                        If dicFieldValueOrder.Item(v) = "" Then
                            WriteLog "缺少字段!"
                            ImportOrderDataToDB = "{'STATE':'ERR','DESC':'表格头部缺少字段:" & CStr(v) & "'}"
                            Exit Function
                        End If

                        strFieldOrder = strFieldOrder & """" & CStr(v) & """," '如果一切正常，顺便把字段顺序也输出了。这样就可以避免因为表格字段顺序不一致，导致数据错乱
                    Next
                    
                    strFieldOrder = strFieldOrder & """ThirdPartExpressNOList"""
                    '代码走到这里，至少表头的字段存在性和顺序都已经做好了，开始做下面的正文内容了。

                Else
                    '检测每一个单元个是否都有合法的值，第一版先写死算了。后续再做外部验证配置文件
                    '检测的同时，顺便就把内容序列化了，方便后期如果合法的话，就直接拿来做导入操作了。
                    '如果不合法的话，也就是把变量直接丢弃了，并给出提示就行了。
                    '问题是批量导入的话，服务器端也要做修改啊~~~~好麻烦的。
                
                    Dim strRecord As String
                    strRecord = ""
                
                    '                For i = 0 To UBound(arrField)
                    '
                    '                Next
                
                    For Each v In dicFieldValueOrder.keys
                
                        'If arrField(dicFieldValueOrder.Item(v)) <> "" Then
                        If VBA.InStr(1, CStr(v), "Date", vbBinaryCompare) > 0 Then
                            arrField(dicFieldValueOrder.Item(v)) = Format(arrField(dicFieldValueOrder.Item(v)), "yyyy-mm-dd")
                        End If

                        Dim strValue As String
                        strValue = modCommon.MakeQueryValue(CStr(v), arrField(dicFieldValueOrder.Item(v)), False)

                        If modCustERR.gERR Then
                        
                            ImportOrderDataToDB = "{'STATE':'ERR','DESC':'" & modCustERR.gERRDESC & "'}"
                            modCustERR.ERRClear
                            Exit Function
                        
                        End If

                        strRecord = strRecord & """" & strValue & ""","
                    
                        'End If

                    Next

                    If dicExpressFieldValue.Item("承运商号码") <> "" Then
                        Dim strVenderCode As String
                        Dim strVenderExpressNOList As String
                        strVenderCode = MappingVenderCode(arrField(dicExpressFieldValue.Item("承运商号码")))
                        strVenderExpressNOList = GetVenderExpressNOList(strVenderCode, arrField(dicExpressFieldValue.Item("运单号码")))
                        '上面是对于一票多件快递的特殊处理，否则快递单号一个格子里有好几个，写进数据库了也没用。
                    
                        'strRecord = strRecord & strVenderCode & "_" & arrField(dicExpressFieldValue.Item("运单号码"))
                        strRecord = strRecord & """" & strVenderExpressNOList & ""","
                    
                        'strRecord = Left(strRecord, Len(strRecord) - 1)
                        strData = strData & strRecord & """|"","
                    Else
                    
                        ImportOrderDataToDB = "{'STATE':'ERR','DESC':'复制的内容不正确!'}"
                        
                        Exit Function
                    
                    End If
                End If
            End If

        Next

        'Debug.Print strFieldOrder
        'Debug.Print strData

        If UBound(Split(strFieldOrder, ",")) > 5 Then '随便定了，就至少要有5列才算有效，因为还有一个是放运单号的，无论是否有运单，先占着位子再说。
            'strFieldOrder = Left(strFieldOrder, Len(strFieldOrder) - 1)
            strData = Left(strData, Len(strData) - 5)
            Dim strPostData As String
        
            strPostData = "{""Type"":""SaveOrderDetail_Batch"",""Fields"":[|-Header-|],""Values"":[|-Records-|]}"
            strPostData = Replace(strPostData, "|-Header-|", strFieldOrder)
            strPostData = Replace(strPostData, "|-Records-|", strData)
            Dim strURL As String
            strURL = "frmorder_detail.asp"
            
            strResult = PostData(strURL, strPostData, 30) '需要设置超时时间长一点，否则大数据量的导入肯定报超时的。
            Debug.Print strResult
            ImportOrderDataToDB = strResult
        Else
        
            ImportOrderDataToDB = "{'STATE':'ERR','DESC':'复制的表格行数异常!'}"
        
        End If

    Else
    
        ImportOrderDataToDB = "{'STATE':'ERR','DESC':'复制的表格行数异常'}"
    
    End If

End Function

Private Function GetVenderExpressNOList(ByVal VenderCode As String, ByVal VenderExpressNOList As String) As String

    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    
    Reg.Global = True
    Reg.IgnoreCase = True
    Reg.MultiLine = False
    
    Reg.Pattern = "([\w\d]{8,})"
    
    Dim Mc As VBScript_RegExp_55.MatchCollection
    
    Dim M As VBScript_RegExp_55.Match
    
    Set Mc = Reg.Execute(VenderExpressNOList)
    
    If Mc.Count > 0 Then
    
        '直接带格式输出算了，不折腾了。以后有空再做抽象。
        Dim strResult As String

        For Each M In Mc
            strResult = strResult & VenderCode & "_" & M.SubMatches(0) & "|"
        Next
        
        GetVenderExpressNOList = Left(strResult, Len(strResult) - 1)
        
    Else
    
        GetVenderExpressNOList = ""
    
    End If
    
    Set M = Nothing
    Set Mc = Nothing
    Set Reg = Nothing

End Function

'Private Function CheckEachFieldValue(ByRef dic As Scripting.Dictionary, ByRef arr() As String) As Boolean
'
'    Dim i As Integer
'
'    For i = 0 To UBound(arr)
'
'        If dic.Exists(arr(i)) Then
'            dic.Item(arr(i)) = 1
'        End If
'
'    Next
'
'    '先把数组里所有的元素都遍历一边，在字典里找对应的KEY，找到就在VALUE里设1
'
'End Function

Public Function GetMappingField(ByVal strCellText As String, ByRef arrOrderFieldHeader() As String, ByRef arrOrderTableFieldHeader() As String) As String
    
    Dim i As Integer
    
    For i = 0 To UBound(arrOrderFieldHeader)
    
        If arrOrderFieldHeader(i) = strCellText Then
        
            GetMappingField = arrOrderTableFieldHeader(i)
            Exit Function
        
        Else
        
        End If
    
    Next

    GetMappingField = ""
End Function

Private Function MappingVenderCode(ByVal strCode As String) As String

    MappingVenderCode = getMappingCode_Core("ThdVender", strCode)

End Function
