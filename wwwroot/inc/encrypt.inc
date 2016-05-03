<%
Public Function EnCode(SourceString)
    Dim ii, s1, nc, nn1, IntTmp ', nn2, s2
    
    Randomize
    nn1 = Int((&H7F) * Rnd + 1)
    'nn2 = Int((&H7FFF) * Rnd + 1)
    s1 = Trim(CStr(nn1)) & "_"
	'S2 = Trim(CStr(nn2)) & "_"
    nn1 = nn1 + 7
    'nn2 = nn2 + 7

    For ii = 1 To Len(SourceString)

        nc = Asc(Mid(SourceString, ii, 1))
		'Response.write nc Xor nn1
		'response.write "|"
		IntTmp = nc Xor nn1
		'If IsNumeric(IntTmp) Then
			If (IntTmp >= 35 And IntTmp <= 47 And IntTmp <> 39) Or (IntTmp >= 64 And IntTmp <= 90) Or (IntTmp >=97 And IntTmp <=122) Then
				'Response.write IntTmp
				'response.write "|"
				s1 = s1 & Chr(IntTmp) & "_"
			Else
				s1 = s1 & Trim(CStr(IntTmp)) & "_"
			End If
		'Else
		'	s1 = s1 & Trim(CStr(IntTmp)) & "_"
		'End If
        
        'S2 = S2 & Trim(CStr(nc Xor nn2)) & "_"

    Next

    'tCodeString = s
    EnCode = Left(S1, Len(S1) -1)'2
End Function

Function Decode(Code)
    Dim ii, jj, S2, nc, nn, arr, i

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
				nc = Cint(CLng(arr(i))) Xor nn
				S2 = S2 & Chr(nc)
			Else
				'If Len(arr(i)) = 1 Then
					'response.write Asc(arr(i))
					'response.write "|"
					nc = Cint(CLng(Asc(arr(i)))) Xor nn
					S2 = S2 & Chr(nc)
				'Else
					'nc = Cint(CLng("&H" & arr(i))) Xor nn
					'S2 = S2 & Asc(nc)
				'End If
			End If
        Next

        Decode = S2

    End If

End Function

%>