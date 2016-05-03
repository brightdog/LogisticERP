Attribute VB_Name = "modImportOdrtoDB"
Option Explicit
Public Const gstrOrderFieldHeader As String = "��װ������|��������������|��ַ|����|�ջ��˵绰|����|ë��|��Ʒ�ܽ��|��ע|ǩ������" 'ǩ����||�����̺���|�˵���"

Public Const gstrOrderTableFieldHeader As String = "ThdPkgNO|ThdOrderOutWorkNO|ReceiverAddress|ReceiverName|ReceiverMobile|PkgNum|PkgWeight|InsurePrice|Remark|ReceiveDateTime"

Public Const gstrExpressFieldHeader As String = "�����̺���|�˵�����"

Public Function ImportOrderDataToDB(ByVal strRawData As String) As String
    Dim i As Integer
    Dim arrLine()  As String
    
    Dim strResult As String
    
    arrLine = Split(strRawData, vbCrLf, -1, vbBinaryCompare)
    
    If UBound(arrLine) > 0 Then '����Ҫ2�У�һ����ͷ��һ������
    
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

                'Ϊÿһ�����ݳ�ʼ���ֵ���󣬼��ÿһ��ֵ�Ƿ��У�������ڵ���֮ǰ������ʾ��
                '����Ҳ���ǵ�һ�е�ʱ����Ҫȷ���ֶ�˳�򣬺���ֻҪ����Ƿ���ֵӦ�þͿ����ˡ�
                If iLine = 0 Then
                    'ȷ���ֶ�˳���Լ��ֶ��Ƿ��д���
                
                    For i = 0 To UBound(arrField)
                        Dim strMappingField As String
                        strMappingField = GetMappingField(arrField(i), arrOrderFieldHeader, arrOrderTableFieldHeader)

                        If strMappingField <> "" Then
                            If dicFieldValue.Exists(strMappingField) Then
                                dicFieldValueOrder.Item(strMappingField) = i '���ÿһ���ֶε�˳��
                            Else
                                'WriteLog "�ж�����ֶ�:" & arrField(i)
                                '����һ�²��������˲���������ˡ�������������ġ�
                            End If
                        End If
                            
                        If dicExpressFieldValue.Exists(arrField(i)) Then
                            dicExpressFieldValue.Item(arrField(i)) = i '���ÿһ���ֶε�˳��
                        Else
                            'WriteLog "�ж�����ֶ�:" & arrField(i)
                            '����һ�²��������˲���������ˡ�������������ġ�
                        End If

                    Next

                    '                For i = 0 To UBound(arrExpressFieldHeader)
                    '
                    '                    dicExpressFieldValue.Add arrExpressFieldHeader(i), ""
                    '
                    '                Next
                    For Each v In dicFieldValueOrder.keys
                
                        If dicFieldValueOrder.Item(v) = "" Then
                            WriteLog "ȱ���ֶ�!"
                            ImportOrderDataToDB = "{'STATE':'ERR','DESC':'���ͷ��ȱ���ֶ�:" & CStr(v) & "'}"
                            Exit Function
                        End If

                        strFieldOrder = strFieldOrder & """" & CStr(v) & """," '���һ��������˳����ֶ�˳��Ҳ����ˡ������Ϳ��Ա�����Ϊ����ֶ�˳��һ�£��������ݴ���
                    Next
                    
                    strFieldOrder = strFieldOrder & """ThirdPartExpressNOList"""
                    '�����ߵ�������ٱ�ͷ���ֶδ����Ժ�˳���Ѿ������ˣ���ʼ����������������ˡ�

                Else
                    '���ÿһ����Ԫ���Ƿ��кϷ���ֵ����һ����д�����ˡ����������ⲿ��֤�����ļ�
                    '����ͬʱ��˳��Ͱ��������л��ˣ������������Ϸ��Ļ�����ֱ����������������ˡ�
                    '������Ϸ��Ļ���Ҳ���ǰѱ���ֱ�Ӷ����ˣ���������ʾ�����ˡ�
                    '��������������Ļ�����������ҲҪ���޸İ�~~~~���鷳�ġ�
                
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

                    If dicExpressFieldValue.Item("�����̺���") <> "" Then
                        Dim strVenderCode As String
                        Dim strVenderExpressNOList As String
                        strVenderCode = MappingVenderCode(arrField(dicExpressFieldValue.Item("�����̺���")))
                        strVenderExpressNOList = GetVenderExpressNOList(strVenderCode, arrField(dicExpressFieldValue.Item("�˵�����")))
                        '�����Ƕ���һƱ�����ݵ����⴦�������ݵ���һ���������кü�����д�����ݿ���Ҳû�á�
                    
                        'strRecord = strRecord & strVenderCode & "_" & arrField(dicExpressFieldValue.Item("�˵�����"))
                        strRecord = strRecord & """" & strVenderExpressNOList & ""","
                    
                        'strRecord = Left(strRecord, Len(strRecord) - 1)
                        strData = strData & strRecord & """|"","
                    Else
                    
                        ImportOrderDataToDB = "{'STATE':'ERR','DESC':'���Ƶ����ݲ���ȷ!'}"
                        
                        Exit Function
                    
                    End If
                End If
            End If

        Next

        'Debug.Print strFieldOrder
        'Debug.Print strData

        If UBound(Split(strFieldOrder, ",")) > 5 Then '��㶨�ˣ�������Ҫ��5�в�����Ч����Ϊ����һ���Ƿ��˵��ŵģ������Ƿ����˵�����ռ��λ����˵��
            'strFieldOrder = Left(strFieldOrder, Len(strFieldOrder) - 1)
            strData = Left(strData, Len(strData) - 5)
            Dim strPostData As String
        
            strPostData = "{""Type"":""SaveOrderDetail_Batch"",""Fields"":[|-Header-|],""Values"":[|-Records-|]}"
            strPostData = Replace(strPostData, "|-Header-|", strFieldOrder)
            strPostData = Replace(strPostData, "|-Records-|", strData)
            Dim strURL As String
            strURL = "frmorder_detail.asp"
            
            strResult = PostData(strURL, strPostData, 30) '��Ҫ���ó�ʱʱ�䳤һ�㣬������������ĵ���϶�����ʱ�ġ�
            Debug.Print strResult
            ImportOrderDataToDB = strResult
        Else
        
            ImportOrderDataToDB = "{'STATE':'ERR','DESC':'���Ƶı�������쳣!'}"
        
        End If

    Else
    
        ImportOrderDataToDB = "{'STATE':'ERR','DESC':'���Ƶı�������쳣'}"
    
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
    
        'ֱ�Ӵ���ʽ������ˣ��������ˡ��Ժ��п���������
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
'    '�Ȱ����������е�Ԫ�ض�����һ�ߣ����ֵ����Ҷ�Ӧ��KEY���ҵ�����VALUE����1
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
