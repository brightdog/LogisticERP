VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncncodeURI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function URLenc(ByVal strEnc As String) As String

Dim lngLoop As Long, lngAsc As Long

Dim strChr As String

For lngLoop = 1 To Len(strEnc)

strChr = Mid(strEnc, lngLoop, 1)

If Abs(Asc(strChr)) < 255 Then

URLenc = URLenc & strChr

Else

lngAsc = Asc(strChr)

If lngAsc < 0 Then lngAsc = lngAsc + 65536

URLenc = URLenc & "%" & Hex((lngAsc And -256) \ 255) & "%" & Hex(lngAsc And 255)

End If

Next

End Function
Public Function UTF8EncodeURI(strInput)
    Dim wch, uch, szRet
    Dim x
    Dim nAsc, nAsc2, nAsc3

    If strInput = "" Then
        UTF8EncodeURI = strInput
        Exit Function
    End If

    For x = 1 To Len(strInput)
        wch = Mid(strInput, x, 1)
        nAsc = AscW(wch)

        If nAsc < 0 Then nAsc = nAsc + 65536

        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else

            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If

    Next

    UTF8EncodeURI = szRet
End Function

Public Function GBKEncodeURI(strInput)
    Dim i As Long
    Dim x() As Byte
    Dim szRet As String

    szRet = ""
    x = StrConv(strInput, vbFromUnicode)

    For i = LBound(x) To UBound(x)
        szRet = szRet & "%" & Hex(x(i))
    Next

    GBKEncodeURI = szRet
End Function

Function UTF2GB(UTFStr)
    Dim Dig As Long
    Dim GBStr As String

    For Dig = 1 To Len(UTFStr)

        If Mid(UTFStr, Dig, 1) = "%" Then
            If Len(UTFStr) >= Dig + 8 Then
                GBStr = GBStr & ConvChinese(Mid(UTFStr, Dig, 9))
                Dig = Dig + 8
            Else
                GBStr = GBStr & Mid(UTFStr, Dig, 1)
            End If

        Else
            GBStr = GBStr & Mid(UTFStr, Dig, 1)
        End If

    Next

    UTF2GB = GBStr
End Function

Private Function ConvChinese(x)
    Dim a() As String
    Dim i As Integer
    Dim j As Integer
    Dim Digs As String
    Dim Unicode As String

    a = Split(Mid(x, 2), "%")
    i = 0
    j = 0
    
    For i = 0 To UBound(a)
        a(i) = c16to2(a(i))
    Next
        
    For i = 0 To UBound(a) - 1
        Digs = InStr(a(i), "0")
        Unicode = ""

        For j = 1 To Digs - 1

            If j = 1 Then
                a(i) = Right(a(i), Len(a(i)) - Digs)
                Unicode = Unicode & a(i)
            Else
                i = i + 1
                a(i) = Right(a(i), Len(a(i)) - 2)
                Unicode = Unicode & a(i)
            End If

        Next
        
        If Len(c2to16(Unicode)) = 4 Then
            ConvChinese = ConvChinese & ChrW(Int("&H" & c2to16(Unicode)))
        Else
            ConvChinese = ConvChinese & Chr(Int("&H" & c2to16(Unicode)))
        End If

    Next

End Function

Private Function c2to16(x)
    Dim i As Integer
    i = 1

    For i = 1 To Len(x) Step 4
        c2to16 = c2to16 & Hex(c2to10(Mid(x, i, 4)))
    Next

End Function
    
Private Function c2to10(x)
    Dim i As Integer
    c2to10 = 0

    If x = "0" Then Exit Function
    i = 0

    For i = 0 To Len(x) - 1

        If Mid(x, Len(x) - i, 1) = "1" Then c2to10 = c2to10 + 2 ^ (i)
    Next

End Function

Private Function c16to2(x)
    Dim i As Integer
    Dim tempstr As String

    i = 0

    For i = 1 To Len(Trim(x))
        tempstr = c10to2(CInt(Int("&h" & Mid(x, i, 1))))

        Do While Len(tempstr) < 4
            tempstr = "0" & tempstr
        Loop

        c16to2 = c16to2 & tempstr
    Next

End Function

Private Function c10to2(x)
    Dim mysign As Integer
    Dim Digs As Integer
    Dim tempnum As Integer
    Dim i As Integer
    mysign = Sgn(x)
    x = Abs(x)
    Digs = 1

    Do

        If x < 2 ^ Digs Then
            Exit Do
        Else
            Digs = Digs + 1
        End If

    Loop

    tempnum = x
    
    i = 0

    For i = Digs To 1 Step -1

        If tempnum >= 2 ^ (i - 1) Then
            tempnum = tempnum - 2 ^ (i - 1)
            c10to2 = c10to2 & "1"
        Else
            c10to2 = c10to2 & "0"
        End If

    Next

    If mysign = -1 Then c10to2 = "-" & c10to2
End Function




Public Function Unicode_Encode(ByVal str As String) As String

   Dim SB As New clsStringBuilder
   Dim i As Long
   Dim j As Long
   Dim aL1 As Variant
   Dim aL2 As Variant
   Dim c As String
   Dim p As Boolean

   aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
   aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
   For i = 1 To Len(str)
    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    Reg.Global = True
    Reg.IgnoreCase = True
    Reg.Pattern = "([^\u4e00-\u9fa5])"
    If Reg.Test(Mid(str, i, 1)) Then
        SB.Append Mid(str, i, 1)
    Else
      p = True
      c = Mid(str, i, 1)
      For j = 0 To 7
         If c = Chr(aL1(j)) Then
            SB.Append "\" & Chr(aL2(j))
            p = False
            Exit For
         End If
      Next

      If p Then
         Dim a
         a = AscW(c)
         If a > 31 And a < 127 Then
            SB.Append c
         ElseIf a > -1 Or a < 65535 Then
            SB.Append "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
         End If
      End If
      End If
   Next
   
   Set Reg = Nothing
   Unicode_Encode = SB.toString
   Set SB = Nothing
   
End Function


Public Function Unicode_Decode_Core(ByVal UnicodeStr As String) As String
    Dim tmpStr() As String
    Dim SB As New clsStringBuilder
    tmpStr = Split(LCase(UnicodeStr), "\u", -1, vbBinaryCompare)
    Dim i As Integer
    For i = 0 To UBound(tmpStr)
    
        If Len(tmpStr(i)) <> 4 Then
            SB.Append ""
        Else
        
            Dim lHex As String
            Dim rHex As String
            rHex = Right(tmpStr(i), 2)
            lHex = Left(tmpStr(i), 2)
            Dim Llong As Long
            Dim Rlong As Long
            Llong = Val("&H" & lHex)
            Rlong = Val("&H" & rHex)
            SB.Append ChrB(Rlong) + ChrB(Llong)
        End If
    Next
    Unicode_Decode_Core = SB.toString
End Function


Public Function Unicode_Decode(ByVal UnicodeStr As String) As String
Dim Reg As New VBScript_RegExp_55.RegExp
Dim Mc As VBScript_RegExp_55.MatchCollection
Dim m As VBScript_RegExp_55.Match

Reg.Global = True
Reg.IgnoreCase = True
Reg.MultiLine = False
Reg.Pattern = "(\\u\w{4})"
Set Mc = Reg.Execute(UnicodeStr)
For Each m In Mc

    UnicodeStr = Replace(UnicodeStr, m.SubMatches(0), Unicode_Decode_Core(m.SubMatches(0)), 1, -1, vbBinaryCompare)

Next
Set Reg = Nothing
Set m = Nothing
Set Mc = Nothing
Unicode_Decode = UnicodeStr
End Function

