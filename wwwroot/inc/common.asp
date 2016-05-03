
<%

Function MakeJsonFromRst(Rst, PageCount, CurrentPage, RsCount, PageSize)
	'全局通用方法， 将数据库返回的记录集打包转换成JSON串，方便网络传输和客户端解析
	'总共5个参数：
	'记录集， 总页数， 当前页数， 总记录数， 分页大小
	
	Dim strJson
	strJson = "{"
	If PageCount > 0 Then
		strJson = strJson & """PageCount"":" & PageCount & ", "
	End If
	If CurrentPage > 0 Then
		strJson = strJson & """CurrentPage"":" & CurrentPage & ", "
	End If
	If RsCount > 0 Then
		strJson = strJson & """RsCount"":" & RsCount & ", "
	End If
	If PageSize > 0 Then
		strJson = strJson & """PageSize"":" & PageSize & ", "
	End If
	'If PageCount > 0 And CurrentPage > 0 And RsCount > 0 And PageSize > 0 Then
	'没办法了，VBSCRIP不支持函数重载（可选参数Optional），只能这么判断一下了，略恶心，不过好歹能用。
		strJson = strJson & """Header"":[""{Header}""], "
	'End If
	strJson = strJson & """Rst"":[""{Rst}""]"
	strJson = strJson & "}"
	
	Dim strHeader
	Dim strRst

	dim i
	i = 0
	Do While Not Rst.Eof
		Dim Fld
		If i = 0 Then
			i = i + 1
			For Each Fld In Rst.Fields
			
				strHeader = strHeader & """" & Fld.Name & ""","
		
			Next
			strHeader = Left(strHeader, Len(strHeader) - 1)
		End If
		Dim strRecord
		strRecord = ""
		For Each Fld In Rst.Fields
			
			strRecord = strRecord & """" & SafeJsonField(Fld.Value) & ""","
		
		Next

		strRst = strRst & "[" & Left(strRecord, Len(strRecord) - 1) & "],"

		Rst.MoveNext
	Loop
	If strRst <> "" Then
		strRst = Left(strRst, Len(strRst) - 1)
	End If
	strJson = Replace(strJson, """{Header}""",strHeader)
	strJson = Replace(strJson, """{Rst}""",strRst)

	MakeJsonFromRst = strJson
End Function


Public Function SafeJsonField(FieldString)

	If FiledString <> "" Then
		FieldString = Replace(FieldString, """", "&quot;")
		FieldString = Replace(FieldString, vbCrLf, "<br>")
		FieldString = Replace(FieldString, ",", "&#44;")
	Else
		FiledString = ""
	End If
	SafeJsonField = FieldString
End Function


Public Function RestoreJsonField(FieldString)
	If FiledString <> "" Then
		FieldString = Replace(FieldString, "&quot;", """")
		FieldString = Replace(FieldString, "<br>", vbCrLf)
		FieldString = Replace(FieldString, "&#44;", ",")
	Else
		FiledString = ""
	End If
	RestoreJsonField = FieldString
End Function

Public Function MakeSqlWhere(F, O, V)

	'Response.write F & vbcrlf
	'Response.write V
	Dim arrField, arrOper, arrValue
	arrField = Split(F, ",")
	arrOper  = Split(O, ",")
	arrValue = Split(V, ",")

	Dim i, strResult
	If Ubound(arrField) = Ubound(arrOper) And Ubound(arrOper) = Ubound(arrValue) Then
		For i = 0 To Ubound(arrField)

			If arrField(i) <> "" And arrValue(i) <> "''" And arrValue(i) <> "" And arrOper(i) <> "" And UCase(arrValue(i)) <> "NULL" Then
				'Response.write arrOper(i) &" = "& arrValue(i)  & vbcrlf
				If arrOper(i) = "LIKE" And InStr(1, arrValue(i), "%") < 1 Then
					arrValue(i) = Left(arrValue(i), Len(arrValue(i)) - 1)  & "%'"
					'Response.write "%%" & vbcrlf
				End If
				strResult = strResult & " And " & arrField(i) & " " & arrOper(i) & " " & arrValue(i) 

			End if

		Next
	Else
		MakeSqlWhere = ""
	End If
	MakeSqlWhere = strResult
End Function

Public Function GetWareHouseIDfromName(strName)

	Dim Rst
	Set Rst = Server.CreateObject("Adodb.RecordSet")

	strSql = "Select WarehouseID From tblWarehouse Where WarehouseName = '" & strName & "'"
	'Response.write strsql & "****"
	Rst.Open strSql, adoConn, 0, 1
	If Not Rst.Eof Then

		GetWareHouseIDfromName = Rst.Fields.Item(0).Value
	Else

		GetWareHouseIDfromName = -1
	End If
	Rst.Close
	Set Rst = Nothing

End Function

































'==================================================================


'函数：FormatDT
'作者：Abo(wupwu@qq.com)
'日期：2008.09.07
'功能：日期时间格式化
'参数：DateTime,日期时间
'　　　Template,格式化模板
'返回：格式化后的字串
'备注：模板标签注释
'　　　yyyy:年
'　　　yy:2位年
'　　　m:月
'　　　mm:补位月，例：01,02
'　　　mmmm:英文月份
'　　　mmm:英文月份缩写
'　　　d:日
'　　　dd:补位日
'　　　h:时
'　　　hh:补位时
'　　　m:分
'　　　mm:补位分
'　　　s:秒
'　　　ss:补位秒
'　　　ww:星期几英文
'　　　w:星期几英文缩写
'修改记录：
'	Wulingfeng(感谢原作者，修改只是为了和VB6的Format函数格式基本一致) @ 2011-07-28
'函数调用示例：      
'Response.Write FormatDT(Now(),"yyyy-mm-dd hh:mm:ss") & "<br>"     
'Response.Write FormatDT(Now(),"yyyy年mm月dd日 hh时mm分ss秒") & "<br>"     
'Response.Write FormatDT(Now(),"w,d mmm yy") & "<br>"     
     
'==================================================================
     
Function FormatDT(DateTime, Template)

    If (Not IsDate(DateTime)) Or Template = "" Then

        FormatDT = Template

        Exit Function

    End If

    Dim dtmY, dtmM, dtmD, dtmH, dtmN, dtmS, dtmW
    Dim arrFW, arrSW, arrFM, arrSM
    dtmY = Year(DateTime)
    dtmM = Month(DateTime)
    dtmD = Day(DateTime)
    dtmH = Hour(DateTime)
    dtmN = Minute(DateTime)
    dtmS = Second(DateTime)
    dtmW = Weekday(DateTime)
    arrFW = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    arrSW = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    arrFM = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    arrSM = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
     
    Template = Replace(Template, "yyyy", dtmY, 1, 1)
    Template = Replace(Template, "yy", Right(dtmY, 2), 1, 1)
    Template = Replace(Template, "mmmm", arrFM(dtmM - 1), 1, 1)
    Template = Replace(Template, "mmm", arrSM(dtmM - 1), 1, 1)

    If InStr(1, Template, "mm", vbBinaryCompare) > 0 Then

        Template = Replace(Template, "mm", Right("00" & dtmM, 2), 1, 1)

    Else

        Template = Replace(Template, "m", dtmM, 1, 1)

    End If


    Template = Replace(Template, "dd", Right("00" & dtmD, 2), 1, 1)
    Template = Replace(Template, "d", dtmD, 1, 1)
    Template = Replace(Template, "hh", Right("00" & dtmH, 2), 1, 1)
    Template = Replace(Template, "h", dtmH, 1, 1)
    Template = Replace(Template, "mm", Right("00" & dtmN, 2), 1, 1)
    Template = Replace(Template, "m", dtmN, 1, 1)
    Template = Replace(Template, "ss", Right("00" & dtmS, 2), 1, 1)
    Template = Replace(Template, "s", dtmS, 1, 1)
    Template = Replace(Template, "ww", arrFW(dtmW - 1), 1, 1)
    Template = Replace(Template, "w", arrSW(dtmW - 1), 1, 1)

    FormatDT = Template

End Function



Function GetDiffDate(StartDate,EndDate)
	Dim h,n,s,strDate
	h = 0
	n = 0
	s = DateDiff("s",StartDate, EndDate)
	strDate = ""

	If s >=3600 Then 
		h = s \ 3600
		s = s - 3600 * h
	End If

	If s >= 60 Then 
		n = s \ 60
		s = s - 60 * n 
	End If 
	strDate = h & ":" 

	If n < 10 Then
		strDate = strDate & "0" & n & ":" 
	Else
		strDate = strDate & n & ":" 
	End If

	If s < 10 Then 
		strDate = strDate & "0" & s
	Else
		strDate = strDate & s
	End If
	GetDiffDate = strDate

End Function

Function FormatTextToHtml(mystr)

	mystr = Replace(mystr & "", vbLf, "<br>")
	mystr = Replace(mystr & "", vbCr, "<br>")
	mystr = Replace(mystr & "", "<br><br>", "<br>")
	mystr = Replace(mystr, " ", "&nbsp;")

	FormatTextToHtml = mystr
End Function

Function FormatHtmlToText(mystr)

	mystr = Replace(mystr & "", "<br>", vbCrLf)
	mystr = Replace(mystr, "&nbsp;", " ")

	FormatTextToHtml = mystr
End Function



'===================================================================================================================



Function FilterCharacter(strWord)
	Dim strTmp
	strTmp = strWord
	strTmp = Replace(strTmp & "", "&", "&amp;")
	strTmp = Replace(strTmp & "", """", "&quot;")


	FilterCharacter = strTmp
End Function




%>


<%
Function aspEncodeURIComponent(sStr)
    aspEncodeURIComponent = myEncodeURIComponent(sStr)
%>
<script language="javascript" type="text/javascript" runat="server">  
  function myEncodeURIComponent(sStr){  
      return encodeURIComponent(sStr);  
  }  
</script>
<%
End Function
%>



