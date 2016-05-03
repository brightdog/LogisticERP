Attribute VB_Name = "JSON"
' VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
' BSD Licensed

Option Explicit

Const INVALID_JSON      As Long = 1
Const INVALID_OBJECT    As Long = 2
Const INVALID_ARRAY     As Long = 3
Const INVALID_BOOLEAN   As Long = 4
Const INVALID_NULL      As Long = 5
Const INVALID_KEY       As Long = 6
Const INVALID_RPC_CALL  As Long = 7

Private bolERR As Boolean

Private psErrors As String

Public Function GetParserErrors() As String
        '<EhHeader>
        On Error GoTo GetParserErrors_Err
        '</EhHeader>
100     GetParserErrors = psErrors
        '<EhFooter>
        Exit Function

GetParserErrors_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.GetParserErrors " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function ClearParserErrors() As String
        '<EhHeader>
        On Error GoTo ClearParserErrors_Err
        '</EhHeader>
100     psErrors = ""
        '<EhFooter>
        Exit Function

ClearParserErrors_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.ClearParserErrors " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

'
'   parse string and create JSON object
'
Public Function Parse(ByRef str As String) As Scripting.Dictionary
        '<EhHeader>
        On Error GoTo Parse_Err
        '</EhHeader>

        Dim index As Long
100     index = 1
102     psErrors = ""
104     bolERR = False
        On Error Resume Next
106     Call skipChar(str, index)
iTry:
108     Debug.Print Mid(str, index, 1)
110     Select Case Mid(str, index, 1)

            Case "{"
112             Set Parse = parseObject(str, index)

114         Case "["
116             Set Parse = parseArray(str, index)

118         Case Else
120             If index < 50 Then
122                 index = index + 1
124                 Call skipChar(str, index)
126                 GoTo iTry
                Else
128                 psErrors = "Invalid JSON"
                End If
        End Select
    
130     If bolERR Then
132         Set Parse = New Scripting.Dictionary
134         Parse.RemoveAll
        End If
    
        '<EhFooter>
        Exit Function

Parse_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.Parse " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

'
'   parse collection of key/value
'
Private Function parseObject(ByRef str As String, ByRef index As Long) As Dictionary
        '<EhHeader>
        On Error GoTo parseObject_Err
        '</EhHeader>

100     Set parseObject = New Dictionary
        Dim sKey As String
   
        ' "{"
102     Call skipChar(str, index)

104     If Mid(str, index, 1) <> "{" Then

106         psErrors = psErrors & "Invalid Object at position " & index & " : " & Mid(str, index) & vbCrLf
            Exit Function

        End If
   
108     index = index + 1

        Do

110         Call skipChar(str, index)

112         If "}" = Mid(str, index, 1) Then

114             index = index + 1
                Exit Do

116         ElseIf "," = Mid(str, index, 1) Then

118             index = index + 1
            
120             Call skipChar(str, index)

            End If
      
            ' add key/value pair
            'Debug.Print "**" & sKey & "**"
122         sKey = parseKey(str, index)
            On Error Resume Next
            '        Debug.Print sKey
            '    If sKey = "RatePlanList" Or sKey = "DayPriceList" Then
            '
            '        Debug.Print sKey
            '        Debug.Print index
            '    End If
    'If sKey = "wi17u000000|206103652" Then
    '        '
    '                Debug.Print sKey
    '        '        Debug.Print index
    '            End If
124         parseObject.Add sKey, parseValue(str, index)

126         If Err.Number <> 0 Or bolERR Then

128             psErrors = psErrors & Err.Description & ": " & sKey & vbCrLf
                Exit Do

            End If

        Loop

eh:

        '<EhFooter>
        Exit Function

parseObject_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.parseObject " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

'
'   parse list
'
Private Function parseArray(ByRef str As String, ByRef index As Long) As Collection
        '<EhHeader>
        On Error GoTo parseArray_Err
        '</EhHeader>

100     Set parseArray = New Collection

        ' "["
102     Call skipChar(str, index)

104     If Mid(str, index, 1) <> "[" Then

106         psErrors = psErrors & "Invalid Array at position " & index & " : " + Mid(str, index, 20) & vbCrLf
            Exit Function

        End If
   
108     index = index + 1

        Do

110         Call skipChar(str, index)

112         If "]" = Mid(str, index, 1) Then

114             index = index + 1
                Exit Do

116         ElseIf "," = Mid(str, index, 1) Then

118             index = index + 1
120             Call skipChar(str, index)

            End If

            ' add value
            On Error Resume Next
122         parseArray.Add parseValue(str, index)

124         If Err.Number <> 0 Or bolERR Then

126             psErrors = psErrors & Err.Description & ": " & Mid(str, index, 20) & vbCrLf
                Exit Do

            End If

        Loop

        '<EhFooter>
        Exit Function

parseArray_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.parseArray " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

'
'   parse string / number / object / array / true / false / null
'
Private Function parseValue(ByRef str As String, ByRef index As Long)
        '<EhHeader>
        On Error GoTo parseValue_Err
        '</EhHeader>

100     Call skipChar(str, index)
        'Debug.Print Mid(str, index, 1)
102     Select Case Mid(str, index, 1)

            Case "{"
                'Debug.Print Mid(str, index, 50)
104             Set parseValue = parseObject(str, index)

106         Case "["
108             Set parseValue = parseArray(str, index)

110         Case """", "'"
112             parseValue = parseString(str, index)

114         Case "t", "f"
116             parseValue = parseBoolean(str, index)

118         Case "n", "u"
120             parseValue = parseNull(str, index)

122         Case Else
124             parseValue = parseNumber(str, index)
        End Select

        '<EhFooter>
        Exit Function

parseValue_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.parseValue " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

'
'   parse string
'
Private Function parseString(ByRef str As String, ByRef index As Long) As String
        '<EhHeader>
        On Error GoTo parseString_Err
        '</EhHeader>

        Dim quote   As String
        Dim Char    As String
        Dim code    As String

        Dim SB As New clsStringBuilder

100     Call skipChar(str, index)
102     quote = Mid(str, index, 1)
104     index = index + 1
   
106     Do While index > 0 And index <= Len(str)

108         Char = Mid(str, index, 1)

110         Select Case (Char)

                Case "\"
112                 index = index + 1
114                 Char = Mid(str, index, 1)

116                 Select Case (Char)

                        Case """", "\", "/", "'"
118                         SB.Append Char
120                         index = index + 1

122                     Case "b"
124                         SB.Append vbBack
126                         index = index + 1

128                     Case "f"
130                         SB.Append vbFormFeed
132                         index = index + 1

134                     Case "n"
136                         SB.Append vbLf
138                         index = index + 1

140                     Case "r"
142                         SB.Append vbCr
144                         index = index + 1

146                     Case "t"
148                         SB.Append vbTab
150                         index = index + 1

152                     Case "u"
154                         index = index + 1
156                         code = Mid(str, index, 4)
158                         SB.Append ChrW(Val("&h" + code))
160                         index = index + 4
                    End Select

162             Case quote
164                 index = index + 1
            
166                 parseString = SB.ToString
168                 Set SB = Nothing
            
                    Exit Function
            
170             Case Else
172                 SB.Append Char
174                 index = index + 1
            End Select

        Loop
   
176     parseString = SB.ToString
178     Set SB = Nothing
   
        '<EhFooter>
        Exit Function

parseString_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.parseString " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

'
'   parse number
'
Private Function parseNumber(ByRef str As String, ByRef index As Long)
        '<EhHeader>
        On Error GoTo parseNumber_Err
        '</EhHeader>

        Dim value   As String
        Dim Char    As String

100     Call skipChar(str, index)

102     Do While index > 0 And index <= Len(str)

104         Char = Mid(str, index, 1)

106         If InStr("+-0123456789.eE", Char) Then

108             value = value & Char
110             index = index + 1

            Else

    '            If InStr(Value, ".") Or InStr(Value, "e") Or InStr(Value, "E") Then
    '
    '                parseNumber = CDbl(Value)
    '
    '            Else

112                 parseNumber = CDbl(value)

    '            End If

                Exit Function

            End If

        Loop

        '<EhFooter>
        Exit Function

parseNumber_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.parseNumber " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

'
'   parse true / false
'
Private Function parseBoolean(ByRef str As String, ByRef index As Long) As Boolean
        '<EhHeader>
        On Error GoTo parseBoolean_Err
        '</EhHeader>

100     Call skipChar(str, index)

102     If Mid(str, index, 4) = "true" Then

104         parseBoolean = True
106         index = index + 4

108     ElseIf Mid(str, index, 5) = "false" Then

110         parseBoolean = False
112         index = index + 5

        Else

114         psErrors = psErrors & "Invalid Boolean at position " & index & " : " & Mid(str, index) & vbCrLf

        End If

        '<EhFooter>
        Exit Function

parseBoolean_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.parseBoolean " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

'
'   parse null
'
Private Function parseNull(ByRef str As String, ByRef index As Long)
        '<EhHeader>
        On Error GoTo parseNull_Err
        '</EhHeader>

100     Call skipChar(str, index)

102     If Mid(str, index, 4) = "null" Then

104         parseNull = Null
106         index = index + 4

        Else

108         psErrors = psErrors & "Invalid null value at position " & index & " : " & Mid(str, index) & vbCrLf

        End If

        '<EhFooter>
        Exit Function

parseNull_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.parseNull " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function parseKey(ByRef str As String, ByRef index As Long) As String
        '<EhHeader>
        On Error GoTo parseKey_Err
        '</EhHeader>

        Dim dquote  As Boolean
        Dim squote  As Boolean
        Dim Char    As String

100     Call skipChar(str, index)

102     Do While index > 0 And index <= Len(str)

104         Char = Mid(str, index, 1)

106         Select Case (Char)

                Case """"
108                 dquote = Not dquote
110                 index = index + 1

112                 If Not dquote Then

114                     Call skipChar(str, index)

116                     If Mid(str, index, 1) <> ":" Then

118                         psErrors = psErrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
                            Exit Do

                        End If

                    End If

120             Case "'"
122                 squote = Not squote
124                 index = index + 1

126                 If Not squote Then

128                     Call skipChar(str, index)

130                     If Mid(str, index, 1) <> ":" Then

132                         psErrors = psErrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
                            Exit Do

                        End If

                    End If

134             Case ":"
136                 index = index + 1

138                 If Not dquote And Not squote Then

                        Exit Do

                    Else

140                     parseKey = parseKey & Char

                    End If

142             Case Else

144                 If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Char) Then

                    Else

146                     parseKey = parseKey & Char

                    End If

148                 index = index + 1
            End Select

        Loop

        '<EhFooter>
        Exit Function

parseKey_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.parseKey " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

'
'   skip special character
'
Private Sub skipChar(ByRef str As String, ByRef index As Long)
        '<EhHeader>
        On Error GoTo skipChar_Err
        '</EhHeader>
        Dim bComment As Boolean
        Dim bStartComment As Boolean
        Dim bLongComment As Boolean

100     Do While index > 0 And index <= Len(str)

102         Select Case Mid(str, index, 1)

                Case vbCr, vbLf

104                 If Not bLongComment Then

106                     bStartComment = False
108                     bComment = False

                    End If
         
110             Case vbTab, " ", "(", ")"
         
112             Case "/"

114                 If Not bLongComment Then

116                     If bStartComment Then

118                         bStartComment = False
120                         bComment = True

                        Else

122                         bStartComment = True
124                         bComment = False
126                         bLongComment = False

                        End If

                    Else

128                     If bStartComment Then

130                         bLongComment = False
132                         bStartComment = False
134                         bComment = False

                        End If

                    End If
         
136             Case "*"

138                 If bStartComment Then

140                     bStartComment = False
142                     bComment = True
144                     bLongComment = True

                    Else

146                     bStartComment = True

                    End If
         
148             Case Else

150                 If Not bComment Then

                        Exit Do

                    End If

            End Select
      
152         index = index + 1

        Loop

        '<EhFooter>
        Exit Sub

skipChar_Err:
        MsgBox Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.skipChar " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function ToString(ByRef obj As Variant) As String
        '<EhHeader>
        On Error GoTo ToString_Err
        '</EhHeader>
        Dim SB As New clsStringBuilder

100     Select Case VarType(obj)

            Case vbNull
102             SB.Append "null"

104         Case vbDate
106             SB.Append """" & CStr(obj) & """"

108         Case vbString
110             SB.Append """" & Encode(obj) & """"

112         Case vbObject
         
                Dim bFI As Boolean
                Dim i As Long
         
114             bFI = True

116             If TypeName(obj) = "Dictionary" Then

118                 SB.Append "{"
                    Dim keys
120                 keys = obj.keys

122                 For i = 0 To obj.Count - 1

124                     If bFI Then bFI = False Else SB.Append ","

                        Dim key
126                     key = keys(i)
128                     SB.Append """" & key & """:" & ToString(obj.Item(key))

130                 Next i

132                 SB.Append "}"

134             ElseIf TypeName(obj) = "Collection" Then

136                 SB.Append "["
                    Dim value

138                 For Each value In obj

140                     If bFI Then bFI = False Else SB.Append ","

142                     SB.Append ToString(value)

144                 Next value

146                 SB.Append "]"

                End If

148         Case vbBoolean

150             If obj Then SB.Append "true" Else SB.Append "false"

152         Case vbVariant, vbArray, vbArray + vbVariant
                Dim sEB
154             SB.Append multiArray(obj, 1, "", sEB)

156         Case Else
158             SB.Append Replace(obj, ",", ".")
        End Select

160     ToString = SB.ToString
162     Set SB = Nothing
   
        '<EhFooter>
        Exit Function

ToString_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.ToString " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function Encode(str) As String
        '<EhHeader>
        On Error GoTo Encode_Err
        '</EhHeader>

        Dim SB As New clsStringBuilder
        Dim i As Long
        Dim j As Long
        Dim aL1 As Variant
        Dim aL2 As Variant
        Dim c As String
        Dim p As Boolean

100     aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
102     aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)

104     For i = 1 To Len(str)

106         p = True
108         c = Mid(str, i, 1)

110         For j = 0 To 7

112             If c = Chr(aL1(j)) Then

114                 SB.Append "\" & Chr(aL2(j))
116                 p = False
                    Exit For

                End If

            Next

118         If p Then

                Dim a
120             a = AscW(c)

122             If a > 31 And a < 127 Then

124                 SB.Append c

126             ElseIf a > -1 Or a < 65535 Then

128                 SB.Append "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)

                End If

            End If

        Next
   
130     Encode = SB.ToString
132     Set SB = Nothing
   
        '<EhFooter>
        Exit Function

Encode_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.Encode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition
        '<EhHeader>
        On Error GoTo multiArray_Err
        '</EhHeader>
   
        Dim iDU As Long
        Dim iDL As Long
        Dim i As Long
   
        On Error Resume Next
100     iDL = LBound(aBD, iBC)
102     iDU = UBound(aBD, iBC)

        Dim SB As New clsStringBuilder

        Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2

104     If Err.Number = 9 Then

106         sPB1 = sPT & sPS

108         For i = 1 To Len(sPB1)

110             If i <> 1 Then sPB2 = sPB2 & ","

112             sPB2 = sPB2 & Mid(sPB1, i, 1)

            Next

            '        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
114         SB.Append ToString(aBD(sPB2))

        Else

116         sPT = sPT & sPS
118         SB.Append "["

120         For i = iDL To iDU

122             SB.Append multiArray(aBD, iBC + 1, i, sPT)

124             If i < iDU Then SB.Append ","

            Next

126         SB.Append "]"
128         sPT = Left(sPT, iBC - 2)

        End If

130     Err.Clear
132     multiArray = SB.ToString
   
134     Set SB = Nothing
        '<EhFooter>
        Exit Function

multiArray_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.multiArray " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

' Miscellaneous JSON functions

Public Function StringToJSON(st As String) As String
        '<EhHeader>
        On Error GoTo StringToJSON_Err
        '</EhHeader>
   
        Const FIELD_SEP = "~"
        Const RECORD_SEP = "|"

        Dim sFlds As String
        Dim sRecs As New clsStringBuilder
        Dim lRecCnt As Long
        Dim lFld As Long
        Dim Fld As Variant
        Dim rows As Variant

100     lRecCnt = 0

102     If st = "" Then

104         StringToJSON = "null"

        Else

106         rows = Split(st, RECORD_SEP)

108         For lRecCnt = LBound(rows) To UBound(rows)

110             sFlds = ""
112             Fld = Split(rows(lRecCnt), FIELD_SEP)

114             For lFld = LBound(Fld) To UBound(Fld) Step 2

116                 sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & Fld(lFld) & """:""" & toUnicode(Fld(lFld + 1) & "") & """")

                Next 'fld

118             sRecs.Append IIf((Trim(sRecs.ToString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"

            Next 'rec

120         StringToJSON = ("( {""Records"": [" & vbCrLf & sRecs.ToString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")

        End If

        '<EhFooter>
        Exit Function

StringToJSON_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.StringToJSON " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
'
'Public Function RStoJSON(rs As ADODB.Recordset) As String
'    On Error GoTo errHandler
'    Dim sFlds As String
'    Dim sRecs As New clsStringBuilder
'    Dim lRecCnt As Long
'    Dim fld As ADODB.Field
'
'    lRecCnt = 0
'
'    If rs.State = adStateClosed Then
'
'        RStoJSON = "null"
'
'    Else
'
'        If rs.EOF Or rs.BOF Then
'
'            RStoJSON = "null"
'
'        Else
'
'            Do While Not rs.EOF And Not rs.BOF
'
'                lRecCnt = lRecCnt + 1
'                sFlds = ""
'
'                For Each fld In rs.Fields
'
'                    sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.Name & """:""" & toUnicode(fld.Value & "") & """")
'
'                Next 'fld
'
'                sRecs.Append IIf((Trim(sRecs.toString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
'                rs.MoveNext
'
'            Loop
'
'            RStoJSON = ("( {""Records"": [" & vbCrLf & sRecs.toString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
'
'        End If
'
'    End If
'
'    Exit Function
'errHandler:
'
'End Function

'Public Function JsonRpcCall(url As String, methName As String, args(), Optional user As String, Optional pwd As String) As Object
'    Dim r As Object
'    Dim cli As Object
'    Dim pText As String
'    Static reqId As Integer
'
'    reqId = reqId + 1
'
'    Set r = CreateObject("Scripting.Dictionary")
'    r("jsonrpc") = "2.0"
'    r("method") = methName
'    r("params") = args
'    r("id") = reqId
'
'    pText = toString(r)
'
'    Set cli = CreateObject("MSXML2.XMLHTTP.6.0")
'   ' Set cli = New MSXML2.XMLHTTP
'    If Len(user) > 0 Then   ' If Not IsMissing(user) Then
'        cli.Open "POST", url, False, user, pwd
'    Else
'        cli.Open "POST", url, False
'    End If
'    cli.setRequestHeader "Content-Type", "application/json"
'    cli.Send pText
'
'    If cli.Status <> 200 Then
'        Err.Raise vbObjectError + INVALID_RPC_CALL + cli.Status, , cli.statusText
'    End If
'
'    Set r = parse(cli.responseText)
'    Set cli = Nothing
'
'    If r("id") <> reqId Then Err.Raise vbObjectError + INVALID_RPC_CALL, , "Bad Response id"
'
'    If r.Exists("error") Or Not r.Exists("result") Then
'        Err.Raise vbObjectError + INVALID_RPC_CALL, , "Json-Rpc Response error: " & r("error")("message")
'    End If
'
'    If Not r.Exists("result") Then Err.Raise vbObjectError + INVALID_RPC_CALL, , "Bad Response, missing result"
'
'    Set JsonRpcCall = r("result")
'End Function

Public Function toUnicode(str As String) As String
        '<EhHeader>
        On Error GoTo toUnicode_Err
        '</EhHeader>

        Dim x As Long
        Dim uStr As New clsStringBuilder
        Dim uChrCode As Integer

100     For x = 1 To Len(str)

102         uChrCode = Asc(Mid(str, x, 1))

104         Select Case uChrCode

                Case 8:   ' backspace
106                 uStr.Append "\b"

108             Case 9: ' tab
110                 uStr.Append "\t"

112             Case 10:  ' line feed
114                 uStr.Append "\n"

116             Case 12:  ' formfeed
118                 uStr.Append "\f"

120             Case 13: ' carriage return
122                 uStr.Append "\r"

124             Case 34: ' quote
126                 uStr.Append "\"""

128             Case 39:  ' apostrophe
130                 uStr.Append "\'"

132             Case 92: ' backslash
134                 uStr.Append "\\"

136             Case 123, 125:  ' "{" and "}"
138                 uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))

140             Case Is < 32, Is > 127: ' non-ascii characters
142                 uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))

144             Case Else
146                 uStr.Append Chr$(uChrCode)
            End Select

        Next

148     toUnicode = uStr.ToString
        Exit Function

        '<EhFooter>
        Exit Function

toUnicode_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.toUnicode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub Class_Initialize()
        '<EhHeader>
        On Error GoTo Class_Initialize_Err
        '</EhHeader>
100     psErrors = ""
102     bolERR = False
        '<EhFooter>
        Exit Sub

Class_Initialize_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot.JSON.Class_Initialize " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

