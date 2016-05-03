VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dynamic Encryption, 每次加密之后的密文都不相同。
Option Explicit

Public Function EnCode(SourceString)
    '2014-08-27结果改成用StringBuilder来拼接，应该可以提高一点速度。
    Dim ii, s1, nc, nn1, IntTmp ', nn2, s2
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    Randomize
    nn1 = Int((&H7F) * Rnd + 1)
    'nn2 = Int((&H7FFF) * Rnd + 1)
    's1 = Trim(CStr(nn1)) & "_"
    SB.Append Trim(CStr(nn1)) & "_"
    'S2 = Trim(CStr(nn2)) & "_"
    nn1 = nn1 + 7
    'nn2 = nn2 + 7
    Dim lngLen As Long
    
    lngLen = Len(SourceString)
    
    For ii = 1 To lngLen

        nc = Asc(mID(SourceString, ii, 1))
        'Response.write nc Xor nn1
        'response.write "|"
        IntTmp = nc Xor nn1

        'If IsNumeric(IntTmp) Then
        If (IntTmp >= 35 And IntTmp <= 47 And IntTmp <> 39) Or (IntTmp >= 64 And IntTmp <= 90) Or (IntTmp >= 97 And IntTmp <= 122) Then
            'Response.write IntTmp
            'response.write "|"
            's1 = s1 & Chr(IntTmp) & "_"
            SB.Append Chr(IntTmp) & "_"
        Else
            's1 = s1 & Trim(CStr(IntTmp)) & "_"
            SB.Append Trim(CStr(IntTmp)) & "_"
        End If

        'Else
        '   s1 = s1 & Trim(CStr(IntTmp)) & "_"
        'End If
        
        'S2 = S2 & Trim(CStr(nc Xor nn2)) & "_"
        DoEvents
    Next

    'tCodeString = s
    s1 = SB.toString
    EnCode = Left(s1, Len(s1) - 1) '2
    Set SB = Nothing
End Function

Function Decode(Code)
    '2014-08-27结果改成用StringBuilder来拼接，应该可以提高一点速度。
    Dim ii, jj, S2, nc, nn, arr, i
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder

    If Code = "" Then
        Decode = ""
    Else

        arr = Split(Code, "_")
        S2 = ""
        
        nn = CLng(arr(0)) + 7

        For i = 1 To UBound(arr)
            'response.write IsNumeric(arr(i))
            
            If IsNumeric(arr(i)) Then
                'arr(i) = CLng("&H" & arr(i))
                nc = CInt(CLng(arr(i))) Xor nn
                'S2 = S2 & Chr(nc)
                SB.Append Chr(nc)
            Else
                'If Len(arr(i)) = 1 Then
                'response.write Asc(arr(i))
                'response.write "|"
                nc = CInt(CLng(Asc(arr(i)))) Xor nn
                'S2 = S2 & Chr(nc)
                SB.Append Chr(nc)
                'Else
                'nc = Cint(CLng("&H" & arr(i))) Xor nn
                'S2 = S2 & Asc(nc)
                'End If
            End If

        Next

        'Decode = S2
        Decode = SB.toString
    End If

    Set SB = Nothing
End Function
