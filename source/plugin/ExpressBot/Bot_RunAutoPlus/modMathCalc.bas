Attribute VB_Name = "modMathCalc"
Option Explicit

Public Function Hex2Dec(InputData As String) As Double
    Dim i As Integer
    Dim decOut As Double
    Dim LenHex As Integer
    Dim HexStep As Double
    Dim MidData As String

    decOut = 0

    InputData = UCase(InputData)
    LenHex = Len(InputData)

    For i = 1 To LenHex
        MidData = Mid(InputData, i, 1)

        If Not (IsNumeric(MidData) Or MidData = "A" Or MidData = "B" _
           Or MidData = "C" Or MidData = "D" Or MidData = "E" Or MidData = "F") Then
            MsgBox "不是合法的十六制进数!请重新输入.", vbExclamation, "数据错误"
            Exit Function
        End If

    Next

    HexStep = 0

    For i = LenHex To 1 Step -1
        HexStep = HexStep * 16

        If HexStep = 0 Then HexStep = 1

        MidData = Mid(InputData, i, 1)

        If MidData = "0" Then
            decOut = decOut + (0 * HexStep)
        ElseIf MidData = "1" Then
            decOut = decOut + (1 * HexStep)
        ElseIf MidData = "2" Then
            decOut = decOut + (2 * HexStep)
        ElseIf MidData = "3" Then
            decOut = decOut + (3 * HexStep)
        ElseIf MidData = "4" Then
            decOut = decOut + (4 * HexStep)
        ElseIf MidData = "5" Then
            decOut = decOut + (5 * HexStep)
        ElseIf MidData = "6" Then
            decOut = decOut + (6 * HexStep)
        ElseIf MidData = "7" Then
            decOut = decOut + (7 * HexStep)
        ElseIf MidData = "8" Then
            decOut = decOut + (8 * HexStep)
        ElseIf MidData = "9" Then
            decOut = decOut + (9 * HexStep)
        ElseIf MidData = "A" Then
            decOut = decOut + (10 * HexStep)
        ElseIf MidData = "B" Then
            decOut = decOut + (11 * HexStep)
        ElseIf MidData = "C" Then
            decOut = decOut + (12 * HexStep)
        ElseIf MidData = "D" Then
            decOut = decOut + (13 * HexStep)
        ElseIf MidData = "E" Then
            decOut = decOut + (14 * HexStep)
        ElseIf MidData = "F" Then
            decOut = decOut + (15 * HexStep)
        End If

    Next

    Hex2Dec = decOut
End Function


Public Function BIN_to_DEC(ByVal Bin As String) As Variant
'******************************************************************
' 用途：将二进制转化为十进制
' 输入：Bin(二进制数)
' 输入数据类型：String
' 输出：BIN_to_DEC(十进制数)
' 输出数据类型：Decimal
' 输入的最大数为(96个1),输出最大数为"79228162514264337593543950335"
' by chiaboyzyq(猴哥) 2010-03-22
'******************************************************************
    Dim i As Integer
    BIN_to_DEC = CDec(BIN_to_DEC)
    For i = 1 To Len(Bin)
        BIN_to_DEC = BIN_to_DEC * 2 + Mid(Bin, i, 1)
    Next i
   
End Function

Public Function DEC_to_BIN(ByVal Dec) As String
    '******************************************************************
    ' 用途：将十进制转化为二进制
    ' 输入：Dec(十进制数)
    ' 输入数据类型：Decimal
    ' 输出：DEC_to_BIN(二进制数)
    ' 输出数据类型：String
    ' 输入的最大数为"79228162514264337593543950335",输出最大数为(96个1)
    ' by chiaboyzyq(猴哥) 2010-03-22
    '******************************************************************

    '668304922822115300
    
    '100101000110010010111100100101010110011101001111111111100100
    '10010111111101000111011101100001101100010111100010
    '10010111111101000111011101100001101100010111100010
    '1011101010100010111110100100000100100000111110110
    Dim Bit As String, tmp As Byte
    DEC_to_BIN = ""

    '    Dec = CDec(Dec)
    Do While Val(Dec) > 0
        tmp = Right(Dec, 1)

        If tmp Mod 2 = 0 Then Bit = 0 Else Bit = 1
        
        DEC_to_BIN = Bit & DEC_to_BIN
        Debug.Print Dec
        Dec = BigDivision(Dec, 2, True, 1)
        
        Debug.Print "BigDivision:" & Dec
        
        Dec = BigADD(Dec, 0.5, False)
        
        Debug.Print "BigADD:" & Dec

        If tmp Mod 2 <> 0 Then
            Dec = BigMinus(Dec, 1)
        End If

    Loop
   
End Function

Public Function AWY(ByVal a, ByVal b) As String
    '原先可能溢出的数据，进行“按位与”运算
    Dim arra() As String
    Dim arrb() As String

    a = DEC_to_BIN(a)
    b = DEC_to_BIN(b)


    Dim offset As Integer

    offset = Len(a) - Len(b)
    Dim i As Integer

    If offset > 0 Then

        For i = 1 To offset
    
            b = "0" & b
    
        Next

    Else

        For i = 1 To offset
    
            a = "0" & a
    
        Next

    End If
    arra = StrToArr(CStr(a))
    arrb = StrToArr(CStr(b))
    Dim strResult As String
    For i = 0 To UBound(arra)
    
        If arra(i) = arrb(i) And arra(i) = 1 Then
            strResult = strResult & "1"
        Else
            strResult = strResult & "0"
        End If
    
    Next
    
    AWY = BIN_to_DEC(strResult)
    
End Function



Public Function AWH(ByVal a, ByVal b) As String
    '原先可能溢出的数据，进行“按位或”运算
    Dim arra() As String
    Dim arrb() As String

    a = DEC_to_BIN(a)
    b = DEC_to_BIN(b)


    Dim offset As Integer

    offset = Len(a) - Len(b)
    Dim i As Integer

    If offset > 0 Then

        For i = 1 To offset
    
            b = "0" & b
    
        Next

    Else

        For i = 1 To offset
    
            a = "0" & a
    
        Next

    End If
    arra = StrToArr(CStr(a))
    arrb = StrToArr(CStr(b))
    Dim strResult As String
    For i = 0 To UBound(arra)
    
        If arra(i) = 1 Or arrb(i) = 1 Then
            strResult = strResult & "1"
        Else
            strResult = strResult & "0"
        End If
    
    Next
    
    AWH = BIN_to_DEC(strResult)
    
End Function
Public Function AWYH(ByVal a, ByVal b) As String
    '原先可能溢出的数据，进行“按位或”运算
    Dim arra() As String
    Dim arrb() As String

    a = DEC_to_BIN(a)
    b = DEC_to_BIN(b)


    Dim offset As Integer

    offset = Len(a) - Len(b)
    Dim i As Integer

    If offset > 0 Then

        For i = 1 To offset
    
            b = "0" & b
    
        Next

    Else

        For i = 1 To offset
    
            a = "0" & a
    
        Next

    End If
    arra = StrToArr(CStr(a))
    arrb = StrToArr(CStr(b))
    Dim strResult As String
    For i = 0 To UBound(arra)
    
        If (arra(i) = 1 Or arrb(i) = 1) And (arra(i) = 1 <> arrb(i)) Then
            strResult = strResult & "1"
        Else
            strResult = strResult & "0"
        End If
    
    Next
    
    AWYH = BIN_to_DEC(strResult)
    
End Function

Private Function StrToArr(ByRef s As String) As String()

    Dim i As Long
    
    Dim L As Long
    
    L = Len(s)
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder

    For i = 1 To L
    
        If i < L Then
            SB.Append Mid(s, i, 1) & Chr(3)
        Else
            SB.Append Mid(s, i, 1)
        End If
    
    Next

    StrToArr = Split(SB.ToString, Chr(3), -1, vbBinaryCompare)

End Function

Public Function BigADD(BJS, JS, Optional ByVal bolNeedDecimal As Boolean = False, Optional ByVal DecimalNum As Integer = 2) As String
    ' Write By Wulf 2014-11-04
    ' 大数相加，按照文本方式进行计算。以免出现VB6自带的恶心科学计数法。
    Dim a() As String
    Dim b() As String
    Dim strResult As String
    Dim strDecimal As String
    Dim i As Long
    a = Split("0" & BJS, ".", 2, vbBinaryCompare)
    b = Split("0" & JS, ".", 2, vbBinaryCompare)
    
    '左边补零，使位数相同
    If Len(a(0)) > Len(b(0)) Then

        For i = 1 To Len(a(0)) - Len(b(0))
            JS = "0" & JS
        Next

    ElseIf Len(b(0)) > Len(a(0)) Then

        For i = 1 To Len(b(0)) - Len(a(0))
            BJS = "0" & BJS
        Next

    Else
        '如果整数部分相同，就不需要补零了。
    End If

    '右边补零，使位数相同
    Select Case True

        Case UBound(a) = 1 And UBound(b) = 1

            If Len(a(1)) > Len(b(1)) Then
    
                For i = 1 To Len(a(1)) - Len(b(1))
                    JS = JS & "0"
                Next
    
            ElseIf Len(b(1)) > Len(a(1)) Then
    
                For i = 1 To Len(b(1)) - Len(a(1))
                    BJS = BJS & "0"
                Next
    
            Else
                '如果小数数部分相同，就不需要补零了。
            End If

        Case UBound(a) = 1 And UBound(b) = 0
            JS = JS & "."

            For i = 1 To Len(a(1))
                JS = JS & "0"
            Next

        Case UBound(a) = 0 And UBound(b) = 1
            BJS = BJS & "."

            For i = 1 To Len(b(1))
                BJS = BJS & "0"
            Next

        Case UBound(a) = 0 And UBound(b) = 0
        
    End Select

    a = StrToArr(CStr(BJS))
    b = StrToArr(CStr(JS))
    Dim intAddNum As Integer
    intAddNum = 0

    For i = UBound(a) To 0 Step -1
    
        If a(i) <> "." Then
            Dim strTmp As String
            
            strTmp = CStr(CInt(a(i)) + CInt(b(i)) + intAddNum)

            If CInt(strTmp) >= 10 Then
                intAddNum = 1
                strResult = Right(strTmp, 1) & strResult
            Else
                intAddNum = 0
                strResult = strTmp & strResult
            End If
            
        Else
            strResult = "." & strResult
        End If
        
    Next
    If intAddNum = 1 Then
        strResult = "1" & strResult
    End If
    a = Split(strResult, ".", 2, vbBinaryCompare)

    If bolNeedDecimal Then

        If UBound(a) = 1 Then
            If DecimalNum > 0 Then
                BigADD = a(0) & "." & Left(a(1), DecimalNum)
            Else
                BigADD = a(0)
            End If
        Else

            If DecimalNum > 0 Then
                BigADD = a(0) & "." & Replace(CStr(10 ^ DecimalNum), "1", "")
            Else
                BigADD = a(0)
            End If
        End If

    Else
        BigADD = a(0)
    End If

End Function

Public Function BigMinus(BJS, JS, Optional ByVal bolNeedDecimal As Boolean = False, Optional ByVal DecimalNum As Integer = 2) As String
    ' Write By Wulf 2014-11-04
    ' 大数相减，按照文本方式进行计算。以免出现VB6自带的恶心科学计数法。
    Dim a() As String
    Dim b() As String
    Dim strResult As String
    Dim strDecimal As String
    Dim strNegativeSign As String
    Dim i As Long
    a = Split("0" & BJS, ".", 2, vbBinaryCompare)
    b = Split("0" & JS, ".", 2, vbBinaryCompare)
    
    '左边补零，使位数相同
    If Len(a(0)) > Len(b(0)) Then

        For i = 1 To Len(a(0)) - Len(b(0))
            JS = "0" & JS
        Next

    ElseIf Len(b(0)) > Len(a(0)) Then

        For i = 1 To Len(b(0)) - Len(a(0))
            BJS = "0" & BJS
        Next

    Else
        '如果整数部分相同，就不需要补零了。
    End If

    '右边补零，使位数相同
    Select Case True

        Case UBound(a) = 1 And UBound(b) = 1

            If Len(a(1)) > Len(b(1)) Then
    
                For i = 1 To Len(a(1)) - Len(b(1))
                    JS = JS & "0"
                Next
    
            ElseIf Len(b(1)) > Len(a(1)) Then
    
                For i = 1 To Len(b(1)) - Len(a(1))
                    BJS = BJS & "0"
                Next
    
            Else
                '如果小数数部分相同，就不需要补零了。
            End If

        Case UBound(a) = 1 And UBound(b) = 0
            JS = JS & "."

            For i = 1 To Len(a(1))
                JS = JS & "0"
            Next

        Case UBound(a) = 0 And UBound(b) = 1
            BJS = BJS & "."

            For i = 1 To Len(b(1))
                BJS = BJS & "0"
            Next

        Case UBound(a) = 0 And UBound(b) = 0
        
    End Select
    
    If Val(BJS) >= Val(JS) Then
        a = StrToArr(CStr(BJS))
        b = StrToArr(CStr(JS))
    Else
        b = StrToArr(CStr(BJS))
        a = StrToArr(CStr(JS))
        strNegativeSign = "-"
    End If

    Dim intMinusNum As Integer
    intMinusNum = 0

    For i = UBound(a) To 0 Step -1
    
        If a(i) <> "." Then
            Dim strTmp As String

            If CInt(a(i)) - intMinusNum - CInt(b(i)) < 0 Then
            
                strTmp = CStr(CInt(a(i)) - intMinusNum + 10 - CInt(b(i)))
                intMinusNum = 1
            Else
            
                strTmp = CStr(CInt(a(i)) - intMinusNum - CInt(b(i)))
                intMinusNum = 0
            End If

            strResult = strTmp & strResult
            
        Else
            strResult = "." & strResult
        End If
        
    Next

    If intMinusNum = 1 Then
        strResult = "1" & strResult
    End If

    a = Split(strResult, ".", 2, vbBinaryCompare)

    If bolNeedDecimal Then

        If UBound(a) = 1 Then
            If DecimalNum > 0 Then
                strResult = a(0) & "." & Left(a(1), DecimalNum)
            Else
                strResult = a(0)
            End If

        Else

            If DecimalNum > 0 Then
                strResult = a(0) & "." & Replace(CStr(10 ^ DecimalNum), "1", "")
            Else
                strResult = a(0)
            End If
        End If

    Else
        strResult = a(0)
    End If

    Do While Left(strResult, 1) = "0"
    
        strResult = Right(strResult, Len(strResult) - 1)
    
    Loop

    If strResult = "" Or Left(strResult, 1) = "." Then
        strResult = "0"
    End If

    BigMinus = strNegativeSign & strResult
End Function


Public Function BigDivision(BCS, CS, Optional ByVal bolNeedDecimal As Boolean = False, Optional ByVal DecimalNum As Integer = 2) As String
    ' Write By Wulf 2014-11-04
    ' 大数相除，按照文本方式进行计算。以免出现VB6自带的恶心科学计数法。
    Dim a() As String

    a = StrToArr(CStr(BCS))

    Dim lenCS As Integer
    lenCS = Len(CS)
    Dim i As Integer
    Dim strMod As String
    Dim strResult As String
    Dim strDecimal As String
    strMod = 0
    Dim intStop As Integer
    intStop = UBound(a)
    '    If bolNeedDecimal Then
    intStop = intStop + DecimalNum
    Dim FindDecimal As Boolean
    Dim bolOver As Boolean

    '    End If
    For i = 0 To intStop

        If i <= UBound(a) Then
            If a(i) = "." Then
                FindDecimal = True
                GoTo NextLoop
            End If
        End If
        
        Dim j As Integer
        Dim CurrentNum As String

        If i > UBound(a) Then
            CurrentNum = CStr(0 + CInt(strMod) * 10)
        Else
            CurrentNum = CStr(Int(a(i)) + CInt(strMod) * 10)
            
        End If

        If i = 0 Then
            If BCS > CS Then

                For j = 1 To lenCS - 1

                    If lenCS <= Len(BCS) Then
                        CurrentNum = CurrentNum & a(j)
                    Else
                        CurrentNum = "0" & CurrentNum
                        bolOver = True
                    End If

                    i = i + 1
                Next

            Else
                bolOver = True
            End If

        End If

        If i <= UBound(a) And Not FindDecimal Then
            strResult = strResult & Split(CStr(CurrentNum / CS), ".", 2, vbBinaryCompare)(0)
            
        Else
            strDecimal = strDecimal & Left(Left(CStr(CurrentNum / CS), 15), 1)
        End If

        'If Not bolOver Then
            strMod = CurrentNum Mod CS
        'End If

NextLoop:

        If bolOver And i > DecimalNum Then
            Exit For
        End If

    Next

    strResult = "0" & strResult

    Do While Left(strResult, 1) = "0"
    
        strResult = Right(strResult, Len(strResult) - 1)
    
    Loop

    If strResult = "" Then
        strResult = "0"
    End If
    a = Split(strResult, ".", 2, vbBinaryCompare)
    If UBound(a) = 1 Then
    If Len(a(1)) > 15 Then
    strResult = Left(strResult, 15)
    End If
    End If
    If bolNeedDecimal Then
        If DecimalNum > 0 Then
            If InStr(1, strResult, ".", vbBinaryCompare) > 0 Then
            
                a = Split(strResult, ".", 2, vbBinaryCompare)
                
                BigDivision = a(0) & "." & Left(a(1), DecimalNum)
            
            Else
                BigDivision = strResult & "." & Left(strDecimal, DecimalNum)
            End If

        Else
            BigDivision = strResult
        End If

    Else
        
        BigDivision = Split(strResult, ".", 2, vbBinaryCompare)(0)
    End If

End Function

Public Function BigAWH(ByVal a, ByVal b) As String
'var w = 4294967296; // 2^32
Dim W
W = 2 ^ 32
Dim aLo As String

aLo = BigMod(a, W)


BigAWH = AWH(aLo, b)
'BigAWH("668304922822115300", "0")
'1450508260==>这里的计算结果
'1450508288==>JS里的计算结果，不知道是为什么啊!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

Public Function BigMod(ByVal BCS As Double, ByVal CS As Double, Optional ByVal bolNeedDecimal As Boolean = False, Optional ByVal DecimalNum As Integer = 0) As String
        ' Write By Wulf 2014-11-04
        ' 大数相除，按照文本方式进行计算。以免出现VB6自带的恶心科学计数法。
        '<EhHeader>
        On Error GoTo BigMod_Err
        '</EhHeader>
        Dim a() As String

100     a = StrToArr(CStr(BCS))

        Dim lenCS As Integer
102     lenCS = Len(CStr(CS))
        Dim i As Integer
        Dim strMod As String
        Dim strResult As String
        Dim strDecimal As String
104     strMod = 0
        Dim intStop As Integer
106     intStop = UBound(a)
        '    If bolNeedDecimal Then
108     intStop = intStop + DecimalNum
        Dim FindDecimal As Boolean
        Dim bolOver As Boolean

        '    End If
110     For i = 0 To intStop

112         If i <= UBound(a) Then
114             If a(i) = "." Then
116                 FindDecimal = True
118                 GoTo NextLoop
                End If
            End If
        
            Dim j As Integer
            Dim CurrentNum As String

120         If i > UBound(a) Then
122             CurrentNum = CStr(0 + CDbl(strMod) * 10)
            Else
124             CurrentNum = CStr(CDbl(a(i)) + CDbl(strMod) * 10)
            
            End If

126         If i = 0 Then
128             If BCS > CS Then

130                 For j = 1 To lenCS - 1

132                     If lenCS <= Len(BCS) Then
134                         CurrentNum = CurrentNum & a(j)
                        Else
136                         CurrentNum = "0" & CurrentNum
138                         bolOver = True
                        End If

140                     i = i + 1
                    Next

                Else
142                 bolOver = True
                End If

            End If
            Dim strCurrentResult As String
144         If i <= UBound(a) And Not FindDecimal Then
146             strCurrentResult = Split(CStr(CurrentNum / CS), ".", 2, vbBinaryCompare)(0)
148             strResult = strResult & strCurrentResult
            
            Else
150             strDecimal = strDecimal & Left(Left(CStr(CurrentNum / CS), 15), 1)
            End If

            'If Not bolOver Then
                'WriteLog CurrentNum & ":" & CS & "%" & strCurrentResult & "*" & i & "/" & intStop
152             strMod = CurrentNum - (CS * IIf(strCurrentResult & "" = "", 0, strCurrentResult))
            'End If

NextLoop:

154         If bolOver And i > DecimalNum Then
                Exit For
            End If

        Next

156     strMod = "0" & strMod

158     Do While Left(strMod, 1) = "0"
    
160         strMod = Right(strMod, Len(strMod) - 1)
    
        Loop

162     If strMod = "" Then
164         strMod = "0"
        End If
 
166         BigMod = strMod


        '<EhFooter>
        Exit Function

BigMod_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot_ADSL.modMathCalc.BigMod " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function DEC_to_HEX(ByVal Dec As String) As String

   Dim a As String
    DEC_to_HEX = ""
    Do While Dec > 0
        a = CStr(BigMod(Dec, 16))
        Select Case a
            Case "10": a = "A"
            Case "11": a = "B"
            Case "12": a = "C"
            Case "13": a = "D"
            Case "14": a = "E"
            Case "15": a = "F"
        End Select
        DEC_to_HEX = a & DEC_to_HEX
        Dec = BigDivision(Dec, 16)
    Loop

End Function

    '============ 位运算 ============
    '位左移
    Public Function SHL(nSource As Long, n As Byte) As Long
        SHL = nSource * 2 ^ n
    End Function
      
    '位右移
    Public Function SHR(nSource As Long, n As Byte) As Long
        SHR = nSource / 2 ^ n
    End Function
      
    '获得指定的位
    Public Function GetBits(nSource As Long, n As Byte) As Boolean
        GetBits = nSource And 2 ^ n
    End Function
      
    '设置指定的位

    Public Function SetBits(nSource As Long, n As Byte) As Long
        SetBits = nSource Or 2 ^ n
    End Function
      
    '清除指定的位
    Public Function ResetBits(nSource As Long, n As Byte) As Long
        ResetBits = nSource And Not 2 ^ n
    End Function


