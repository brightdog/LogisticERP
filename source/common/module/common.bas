Attribute VB_Name = "modCommon"
Option Explicit

Public Function convertCRLF(ByRef Value As String) As String
    convertCRLF = Replace(Value, vbCr, Chr$(3))
    convertCRLF = Replace(convertCRLF, vbLf, Chr$(4))
End Function

Public Function restoreCRLF(ByRef Value As String) As String
    restoreCRLF = Replace(Value, Chr$(3), vbCr)
    restoreCRLF = Replace(restoreCRLF, Chr$(4), vbLf)
End Function

Public Function MakeSearchParam(ByRef OwmerForm As VB.Form) As Scripting.Dictionary

    Dim dicParam As Scripting.Dictionary
    Set dicParam = New Scripting.Dictionary

    Dim ctl As VB.Control

    For Each ctl In OwmerForm.Controls
        
        Select Case TypeName(ctl)
            
            Case "TextBox", "ComboBox"
                dicParam.Add ctl.name, ctl.Text

            Case "CheckBox"
                dicParam.Add ctl.name, ctl.Value
        End Select
    
    Next

    Set MakeSearchParam = dicParam

End Function


Public Function CollectionToString(ByRef col As VBA.Collection, Optional ByVal strSplit As String = vbTab) As String

    Dim i As Long
    Dim strResult As String
    
    For i = 1 To col.Count
    
        strResult = strResult & col.Item(i) & strSplit
    
    Next

    CollectionToString = strResult
End Function

Public Function LoadNaviLeftIcon(ByRef obj As MSComctlLib.ListView) As String

    obj.ListItems.Add , "ORDER", "订单"

End Function

Public Function CheckLogin(ByVal strUserName As String, ByVal strPassword As String) As String

    Dim iWeb As clsXMLHTTPGetHtml
    Set iWeb = New clsXMLHTTPGetHtml
    Dim objTransPRD As clsTransformPWD
    Set objTransPRD = New clsTransformPWD
    
    iWeb.URL = gHTTPURL & "chklogin.asp"
    iWeb.PostData = "u=" & strUserName & "&p=" & objTransPRD.TransFormPWD(strPassword)
    
    Call iWeb.Send
    
    CheckLogin = iWeb.ReturnData
    Set objTransPRD = Nothing
    Set iWeb = Nothing
End Function


Public Function convertColtoArray(ByRef col As VBA.Collection) As String()

    Dim i As Long
    
    Dim arr() As String
    
    ReDim arr(0)
    
    For i = 1 To col.Count

        If i > 1 Then
            ReDim Preserve arr(UBound(arr) + 1)
        End If
        
        arr(UBound(arr)) = col.Item(i)
    
    Next
    
    convertColtoArray = arr
    
End Function

Public Function MakeQueryField(ByVal controlName As String) As String
    
    Dim strResult As String

    If Left(controlName, 3) = "txt" Then
        strResult = Right(controlName, Len(controlName) - 3)
    ElseIf Left(controlName, 3) = "cbo" Then
        strResult = Right(controlName, Len(controlName) - 3)
    ElseIf Left(controlName, 3) = "chk" Then
        strResult = Right(controlName, Len(controlName) - 3)
    Else
        strResult = controlName
    End If

    Select Case True
    
        Case Right(controlName, 5) = "_From"
        
            strResult = Left(strResult, Len(strResult) - 5)

        Case Right(controlName, 3) = "_To"
        
            strResult = Left(strResult, Len(strResult) - 3)
            
        Case Right(controlName, 8) = "_BolFrom"
        
            strResult = Left(strResult, Len(strResult) - 8)
            
        Case Else
            'strResult = strResult
    End Select

    MakeQueryField = strResult

End Function

Public Function MakeQueryOperSymbol(ByVal controlName As String, ByVal controlText As String) As String
    
    Dim strResult As String

    If Left(controlName, 3) = "txt" Then
        strResult = Right(controlName, Len(controlName) - 3)
    Else
        strResult = controlName
    End If

    Select Case True
    
        Case Right(controlName, 5) = "_From"
        
            strResult = ">="
        
        Case Right(controlName, 3) = "_To"
        
            strResult = "<="

        Case InStr(1, controlText, "%", vbBinaryCompare) > 0
            strResult = "LIKE"

        Case InStr(1, VBA.LCase(controlName), "name", vbBinaryCompare) > 0
            strResult = "LIKE"

        Case Right(controlName, 8) = "_BolFrom"

            If controlText Then
                strResult = ">="
            Else
                strResult = ""
            End If

        Case Else
            strResult = "="
    End Select

    MakeQueryOperSymbol = strResult

End Function

Public Function MakeQueryValue(ByVal FieldName As String, ByVal controlText As String, Optional NeedQuote As Boolean = True) As String
    '这里需要用到先前modMain里定义的全局数据库对象gdicDBConfig
    '用来判断字段类型，确定是否需要加单引号
    Dim strResult As String
    Dim colFieldInfo As VBA.Collection
    
    Set colFieldInfo = FindFieldInfo(FieldName)

    If Not colFieldInfo Is Nothing Then

        Select Case colFieldInfo(2)
        
            Case "datetime"

                If IsDate(controlText) Then
                    If NeedQuote Then
                        strResult = "'" & JSON.SafeJsonField(controlText) & "'"
                    Else
                        strResult = JSON.SafeJsonField(controlText)
                    End If

                Else
                    strResult = ""
                End If

            Case "varchar"

                If Len(controlText) > (FindFieldInfo(FieldName)(3)) Then
                    modCustERR.gERR = True
                    modCustERR.gERRDESC = "输入超长：" & FieldName & ":" & controlText & "#最长:" & FindFieldInfo(FieldName)(3)
                Else

                    If NeedQuote Then
                        strResult = "'" & JSON.SafeJsonField(controlText) & "'"
                    Else
                        strResult = JSON.SafeJsonField(controlText)
                    End If
                End If

            Case "nvarchar"

                If Len(controlText) > (FindFieldInfo(FieldName)(3) / 2) Then
                    modCustERR.gERR = True
                    modCustERR.gERRDESC = "输入超长：" & FieldName & ":" & controlText & "#最长:" & FindFieldInfo(FieldName)(3) / 2
                Else

                    If NeedQuote Then
                        strResult = "'" & JSON.SafeJsonField(controlText) & "'"
                    Else
                        strResult = JSON.SafeJsonField(controlText)
                    End If
                End If

            Case "date"

                If IsDate(controlText) Then
                    If NeedQuote Then
                        strResult = "'" & JSON.SafeJsonField(controlText) & "'"
                    Else
                        strResult = JSON.SafeJsonField(controlText)
                    End If
                End If

            Case "int"
                
                If Not IsNumeric(controlText) Then
                    modCustERR.gERR = True
                    modCustERR.gERRDESC = "输入类型不匹配：""" & FieldName & """:""" & controlText & """, #应该是整数形式:" & FindFieldInfo(FieldName)(2)
                Else
            
                strResult = controlText
                End If
            Case "bit"
                strResult = controlText
            Case Else
                strResult = controlText
        End Select
    Else
        'strResult = controlText
    End If

    MakeQueryValue = strResult

End Function

Public Function FindFieldInfo(ByVal FieldName As String) As VBA.Collection

    Dim v As Variant
    Dim i As Integer
    Dim bolFind As Boolean

    For Each v In gdicDBConfig.keys
        For i = 1 To gdicDBConfig.Item(v).Item("Rst").Count
            
            If FieldName = gdicDBConfig.Item(v).Item("Rst")(i)(1) Then
                Set FindFieldInfo = gdicDBConfig.Item(v).Item("Rst")(i)
                bolFind = True
        
                Exit For
            End If

        Next

        If bolFind Then
            Exit For
        End If

    Next
    
    If Not bolFind Then
    
        Debug.Print "Internal ERR:" & vbCrLf & "FieldInfo Not Found @ " & FieldName
    
    End If
    
End Function

Public Function isCtlLinkedDB(ByRef ctl As VB.Control) As Boolean

    If (TypeName(ctl) = "TextBox" Or TypeName(ctl) = "ComboBox" Or Left(TypeName(ctl), 3) = "TDB" Or TypeName(ctl) = "CheckBox") And ctl.name <> "txtCreateDT" Then
        isCtlLinkedDB = True
    Else
        isCtlLinkedDB = False
    End If

End Function

Public Function SearchPagedList(ByVal frmName As String, ByRef dicSearchParam As Scripting.Dictionary, Optional ByVal PageSize As Integer = 0, Optional ByVal PageNum As Long = 1) As Scripting.Dictionary

    Dim v As Variant
    Dim SBField As clsStringBuilder
    Set SBField = New clsStringBuilder
    Dim SBValue As clsStringBuilder
    Set SBValue = New clsStringBuilder
    Dim SBOper As clsStringBuilder
    Set SBOper = New clsStringBuilder

    For Each v In dicSearchParam.keys
    
        If CStr(v) <> "AddtionalQueryString" Then

            '如果是日期型的或者带有范围的，那必须得经过一层转换（查询条件mapping）之后，才能进行参数传递。
            '原则上是客户端直接处理好了之后，进行发送。
            '服务器不负责处理具体的查询条件逻辑。
            '减轻服务器上的逻辑处理压力，便于后期维护。
            '例子：字段名称中带有：_Form; _To 的字段，根据前缀，进行处理
            If dicSearchParam.Item(v) & "" <> "" Then
                Dim strDBField As String
                Dim strOperSymbol As String
                Dim strValue As String
                Dim strFieldName As String
                strFieldName = CStr(v)
                strDBField = MakeQueryField(strFieldName)
                strOperSymbol = MakeQueryOperSymbol(strFieldName, dicSearchParam.Item(v))
                If strOperSymbol = "<=" And VBA.IsDate(dicSearchParam.Item(v)) Then
                    strValue = MakeQueryValue(strDBField, DateAdd("d", 1, dicSearchParam.Item(v)))
                Else
                strValue = MakeQueryValue(strDBField, dicSearchParam.Item(v))
                End If
                SBField.Append """" & strDBField & ""","
                SBOper.Append """" & strOperSymbol & ""","
                SBValue.Append """" & strValue & ""","
            End If
        End If

    Next
    
    Dim strAddtionalQueryString As String
    Dim strPostData As String
    Dim strURL As String
    Dim strFields As String
    Dim strValues As String
    Dim strOpers As String
    strFields = SBField.toString
    strOpers = SBOper.toString
    strValues = SBValue.toString
    
    If strFields <> "" Then
        strFields = Left(strFields, Len(strFields) - 1)
    End If
    
    If strOpers <> "" Then
        strOpers = Left(strOpers, Len(strOpers) - 1)
    End If
    
    If strValues <> "" Then
        strValues = Left(strValues, Len(strValues) - 1)
    End If
    
    strAddtionalQueryString = dicSearchParam.Item("AddtionalQueryString")
    
    strPostData = "{""Type"":""PagedList"",""Fields"":[" & strFields & "],""Opers"":[" & strOpers & "],""Values"":[" & strValues & "], ""AddtionalQueryString"":""" & strAddtionalQueryString & """, ""PageSize"":""" & PageSize & """, ""PageNum"":""" & PageNum & """}"
    strURL = LCase(frmName) & ".asp"
    Dim strResult As String
    strResult = PostData(strURL, strPostData)
    Debug.Print strPostData & vbCrLf & "'====================" & vbCrLf & strResult
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)

    Set SearchPagedList = dicResult
End Function

Public Function IncrementalSearch(ByVal TableName As String, ByVal strField As String, ByVal strOper As String, ByVal strValue As String) As Scripting.Dictionary
    '直接返回一个数组，到时候丢到用户控件里去。
    
    Dim strPostData As String
    Dim strURL As String
    
    Dim PageSize As Integer
    PageSize = 20
        
    strPostData = "{"
    strPostData = strPostData & """Type"":""IncrementalList"""
    strPostData = strPostData & ", ""PageSize"":" & PageSize & ""
    strPostData = strPostData & ", ""TableName"":""" & TableName & """"
    strPostData = strPostData & ",""strField"":""" & strField & """"
    strPostData = strPostData & ",""Oper"":""" & strOper & """"

    Select Case LCase(strOper)

        Case "like"
            strPostData = strPostData & ",""Value"":""%" & strValue & "%"""

        Case "="
            strPostData = strPostData & ",""Value"":""" & strValue & """"

        Case Else
            '好吧，默认自己的代码不会写错，就这样了。。。
    End Select

    strPostData = strPostData & "}"
    strURL = "incrementalsearch.asp"
    Dim strResult As String
    strResult = PostData(strURL, strPostData)
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)

    Set IncrementalSearch = dicResult

End Function

Public Function SearchCustInfoByCode(ByVal CustCode As String) As Scripting.Dictionary
    
    Set SearchCustInfoByCode = SearchSimpleList("vwCust_Simple", " And CustCode = '" & JSON.SafeJsonField(CustCode) & "'", "")

End Function
Public Function SearchCustInfoByName(ByVal CustName As String, Optional ByVal strTable As String = "vwCust_Simple", Optional ByVal strColumn As String = "CustName") As Scripting.Dictionary
    
    Set SearchCustInfoByName = SearchSimpleList(strTable, " And " & strColumn & " = '" & JSON.SafeJsonField(CustName) & "'", "")

End Function
Public Function FindDicKeyByValue(ByVal txt As String, ByRef dic As Scripting.Dictionary) As String

    Dim v As Variant
    FindDicKeyByValue = ""

    For Each v In dic.keys
    
        If dic.Item(v) = txt Then
            FindDicKeyByValue = CStr(v)
            Exit For
        End If
    
    Next

End Function

Public Function FindSubAreaByFather(ByVal strFatherKey As String) As Scripting.Dictionary

    Dim dic As Scripting.Dictionary
    
    Dim v As Variant
    
    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    
    Reg.Global = False
    Reg.IgnoreCase = False
    Reg.MultiLine = True
    Reg.Pattern = "," & strFatherKey & "$"
    
    For Each v In gdicLocation.keys

        If Reg.Test(CStr(v)) Then
        
            Set dic = gdicLocation.Item(v)
            Exit For
        End If
        
    Next
    
    Set FindSubAreaByFather = dic
End Function

Public Function SearchSimpleList(ByVal TableName As String, ByVal QueryString As String, Optional ByVal OrderBy As String = "Order By CreateDT Desc") As Scripting.Dictionary
    Dim strPostData, strURL As String
    '_getsamplelistbyid
    strPostData = "{""Type"":""SimpleList"",""TableName"":""" & TableName & """,""QueryString"":""" & SafeJsonField(QueryString) & """,""OrderBy"":""" & SafeJsonField(OrderBy) & """}"
    strURL = "_getsamplelist.asp"
    Dim strResult As String
    strResult = PostData(strURL, strPostData)
    Debug.Print strPostData & vbCrLf & "'====================" & vbCrLf & strResult
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)

    Set SearchSimpleList = dicResult
End Function

'============================== Input Pickup Receipt ===================================
Public Function CheckPickupReceiptInfoByID(ByVal strNo As String) As Scripting.Dictionary

    Dim strResult As String
    strResult = PostData("checkpickupreceiptinfobyid.asp", "{""id"":""" & strNo & """}")  '才一个参数，也不能偷懒！！
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)
    
    Set CheckPickupReceiptInfoByID = dicResult

End Function

Public Function loadCurrentPickupReceiptsExpressNoList(ByVal strNo As String) As Scripting.Dictionary
        '<EhHeader>
        On Error GoTo loadCurrentPickupReceiptsOrderList_Err
        '</EhHeader>
        Dim dicResult As Scripting.Dictionary
100     'Set dicResult = New Scripting.Dictionary
    
        Dim strResult As String
    
102     strResult = PostData("loadpickupreceiptsexpressnolist.asp", "{""id"":""" & strNo & """}") '才一个参数，也不能偷懒！！
        
104     Set dicResult = JSON.Parse(strResult)
    
        Set loadCurrentPickupReceiptsExpressNoList = dicResult
    
        '<EhFooter>
        Exit Function

loadCurrentPickupReceiptsOrderList_Err:
        WriteLog Err.Description & vbCrLf & _
           "in LogisticERP.modCommon.loadCurrentPickupReceiptsExpressNoList " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
'==============================/ Input Pickup Receipt ===================================

'============================== Input OutWarehouse Receipt ===================================
'Public Function CheckOutWarehouseReceiptNOExist(ByVal strNo As String) As Boolean
'
'    Dim strResult As String
'    strResult = PostData("loadoutwarehousereceiptsexpressnolist.asp", "id=" & strNo) '才一个参数，还是简化一点吧，这右里偷懒了！！
'    If strResult = "OK" Then
'
'        CheckOutWarehouseReceiptNOExist = True
'
'    Else
'        CheckOutWarehouseReceiptNOExist = False
'    End If
'
'
'End Function

Public Function CheckOutWarehouseReceiptInfoByID(ByVal strNo As String) As Scripting.Dictionary

    Dim strResult As String
    strResult = PostData("checkoutwarehousereceiptinfobyid.asp", "{""id"":""" & strNo & """}")  '才一个参数，也不能偷懒！！
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)
    
    Set CheckOutWarehouseReceiptInfoByID = dicResult

End Function

Public Function loadCurrentOutWarehouseReceiptsExpressNoList(ByVal strNo As String) As Scripting.Dictionary
        '<EhHeader>
        On Error GoTo loadCurrentOutWarehouseReceiptsExpressNoList_Err
        '</EhHeader>
        
        If Trim(strNo) <> "" Then
            Dim dicResult As Scripting.Dictionary
100         'Set dicResult = New Scripting.Dictionary
        
            Dim strResult As String
        
102         strResult = PostData("loadoutwarehousereceiptsexpressnolist.asp", "{""id"":""" & strNo & """}") '才一个参数，还是简化一点吧，这继续里偷懒！！
            
104         Set dicResult = JSON.Parse(strResult)
        
            Set loadCurrentOutWarehouseReceiptsExpressNoList = dicResult
        End If
        '<EhFooter>
        Exit Function

loadCurrentOutWarehouseReceiptsExpressNoList_Err:
        WriteLog Err.Description & vbCrLf & _
           "in LogisticERP.modCommon.loadCurrentOutWarehouseReceiptsExpressNoList " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

'============================== Input InWarehouse Receipt ===================================

Public Function loadCurrentInWarehouseReceiptsExpressNoList(ByVal strNo As String) As Scripting.Dictionary
        '<EhHeader>
        On Error GoTo loadCurrentInWarehouseReceiptsExpressNoList_Err
        '</EhHeader>
        Dim dicResult As Scripting.Dictionary
100     'Set dicResult = New Scripting.Dictionary
    
        Dim strResult As String
    
102     strResult = PostData("loadinwarehousereceiptsexpressnolist.asp", "{""id"":""" & strNo & """}") '才一个参数，还是简化一点吧，这继续里偷懒！！
        
104     Set dicResult = JSON.Parse(strResult)
    
        Set loadCurrentInWarehouseReceiptsExpressNoList = dicResult
    
        '<EhFooter>
        Exit Function

loadCurrentInWarehouseReceiptsExpressNoList_Err:
        WriteLog Err.Description & vbCrLf & _
           "in LogisticERP.modCommon.loadCurrentInWarehouseReceiptsExpressNoList " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function GetWarehouseListbyEmpName(Optional ByVal USERNAME As String = "") As Scripting.Dictionary
    Set GetWarehouseListbyEmpName = SearchSimpleList("vwWarehouseListbyEmpID", " And EmpID='" & gUSERNAME & "'", "")
End Function

Public Function GetWarehouseListbyID(Optional ByVal ID As String = "") As Scripting.Dictionary

    If ID <> "" Then
        Set GetWarehouseListbyID = SearchSimpleList("vwWarehouseListbyEmpID", " And WarehouseID=" & ID, "")
    Else
        Set GetWarehouseListbyID = SearchSimpleList("vwWarehouseListbyEmpID", "", "")
    End If

End Function

'============================== Input PackageDelivery Receipt ===================================
Public Function CheckPackageDeliveryReceiptInfoByID(ByVal strNo As String) As Scripting.Dictionary

    Dim strResult As String
    strResult = PostData("checkpackagedeliveryreceiptinfobyid.asp", "{""id"":""" & strNo & """}")  '才一个参数，也不能偷懒！！
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)
    
    Set CheckPackageDeliveryReceiptInfoByID = dicResult

End Function

Public Function loadCurrentPackageDeliveryReceiptsExpressNoList(ByVal strNo As String) As Scripting.Dictionary
        '<EhHeader>
        On Error GoTo loadCurrentPackageDeliveryReceiptsExpressNoList_Err
        '</EhHeader>
        Dim dicResult As Scripting.Dictionary
100     'Set dicResult = New Scripting.Dictionary
    
        Dim strResult As String
    
102     strResult = PostData("loadpackagedeliveryreceiptsexpressnolist.asp", "{""id"":""" & strNo & """}") '才一个参数，还是简化一点吧，这继续里偷懒！！
        
104     Set dicResult = JSON.Parse(strResult)
    
        Set loadCurrentPackageDeliveryReceiptsExpressNoList = dicResult
    
        '<EhFooter>
        Exit Function

loadCurrentPackageDeliveryReceiptsExpressNoList_Err:
        WriteLog Err.Description & vbCrLf & _
           "in LogisticERP.modCommon.loadCurrentPackageDeliveryReceiptsExpressNoList " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
'==============================/ Input PackageDelivery Receipt ===================================

Public Function GetWareHouseIDByNameWithSampleDic(ByRef dic As Scripting.Dictionary, ByVal strName As String) As String
    '{"Header":["WarehouseID","WarehouseName","EmpID","EngName"], "Rst":[["1","上海仓","1","admin"],["2","北京仓","1","admin"]]}
    GetWareHouseIDByNameWithSampleDic = 0
    Dim v As Variant
    Dim i As Integer

    For i = 1 To dic.Item("Rst").Item(1).Count
    
        If dic.Item("Rst").Item(i).Item(2) = strName Then
            GetWareHouseIDByNameWithSampleDic = dic.Item("Rst").Item(i).Item(1)
            Exit Function
        End If

    Next

End Function

Public Function doMappingHeadTitle(ByVal strTitle As String) As String
    
    If gdicTitleMapping.Exists(strTitle) Then
    
        strTitle = gdicTitleMapping.Item(strTitle)
    
    Else
    
    End If
    
    doMappingHeadTitle = strTitle
End Function

Public Function GetTitleMapping(ByVal strMapping As String) As Scripting.Dictionary

    Dim dicResult As Scripting.Dictionary
    Set dicResult = New Scripting.Dictionary
    
    Dim arrLine() As String
    arrLine = Split(strMapping, vbCrLf, -1, vbBinaryCompare)
    
    If UBound(arrLine) > -1 Then
    
        Dim i As Integer
        
        For i = 0 To UBound(arrLine)
            
            If arrLine(i) <> "" Then
            
                Dim arrField() As String
                arrField = Split(arrLine(i), vbTab, 2, vbBinaryCompare)
                
                If UBound(arrField) = 1 Then
                
                    If Not dicResult.Exists(arrField(0)) Then
                    
                        dicResult.Add arrField(0), arrField(1)
                    
                    End If
                
                
                End If
            
            
            
            End If
        
        
        Next
    
    
    
    
    Else
    
        WriteLog "*Mapping File is Empty"
    
    End If
    Set GetTitleMapping = dicResult
End Function

Public Function GetOrderIDFromGrid(ByRef GRD As MSFlexGridLib.MSFlexGrid, ByVal ColPos As Integer) As String()

        'me.grdList.TextMatrix(1,0) ::(x,y) x从1开始，因为0表示列头，y从0开始，往右数
        Dim i As Integer
    
        Dim iRows As Integer
    
100     iRows = GRD.rows - 1
        Dim strResult As String

102     For i = 1 To iRows
        
104         strResult = strResult & GRD.TextMatrix(i, ColPos)

106         If i < iRows Then
108             strResult = strResult & "|-|"
            End If
    
        Next

        Dim arrResult() As String

110     If strResult <> "" Then
    
112         arrResult = Split(strResult, "|-|", -1, vbBinaryCompare)

        End If

114     GetOrderIDFromGrid = arrResult

End Function

Public Function GetExpressNOList(ByRef GRD As MSFlexGridLib.MSFlexGrid, ByVal ColPos As Integer) As String()
    
        Dim arrOrderID() As String
100     arrOrderID = GetOrderIDFromGrid(GRD, ColPos)
        Dim arrResult() As String
        Dim i As Integer
        Dim strTmp As String '用来梳理一下orderid用，防止里面有空元素，造成Sql执行报错

102     For i = 0 To UBound(arrOrderID)
        
104         If arrOrderID(i) <> "" Then
        
106             strTmp = strTmp & arrOrderID(i) & ","
        
            Else
        
            End If
    
        Next
    
108     If strTmp <> "" Then
    
110         strTmp = Left(strTmp, Len(strTmp) - 1)
        
            Dim strResult As String
            Dim dicResult As Scripting.Dictionary
    
112         strResult = PostData("getexpressnolistbyorderid.asp", "{""OrderID"":""" & strTmp & """}") '才一个参数，还是简化一点吧，这继续里偷懒！！
        
114         Set dicResult = JSON.Parse(strResult)
        
116         If dicResult.Exists("Rst") Then
        
118             If dicResult.Item("Rst").Count > 0 Then

                    
120                 ReDim arrResult(dicResult.Item("Rst").Count - 1)

122                 For i = 1 To dicResult.Item("Rst").Count
                        Dim Tmp As VBA.Collection
124                     Set Tmp = dicResult.Item("Rst").Item(i)
126                     arrResult(i - 1) = Tmp(1) & "_" & Tmp(2)
                
                    Next
            
                End If
        
            Else
        
            End If
    
        Else
    
        End If

128     GetExpressNOList = arrResult
End Function

Public Function GetOrderIDFromDic(ByVal dicResult As Scripting.Dictionary) As String

    If dicResult.Exists("Header") Then
        If dicResult.Item("Rst")(1).Count > 0 Then
            Dim i As Integer

            For i = 1 To dicResult.Item("Header").Count

                If dicResult.Item("Header").Item(i) = "OrderID" Then
                    
                    GetOrderIDFromDic = dicResult.Item("Rst")(1).Item(i)
                    Exit Function
                End If

            Next
    
        Else
            GetOrderIDFromDic = 0
        End If

    Else
        GetOrderIDFromDic = 0
    End If

End Function

