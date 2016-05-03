<%
Function DeTransform(strPWD)
    Dim i
    i = 0
    Dim sb 
	sb = ""
    For i = 0 To Len(strPWD) - 1

        Select Case i Mod 6

            Case 0:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) + 8)

            Case 1:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) + 3)
     
            Case 2:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) + 9)

            Case 3:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) + 6)
         
            Case 4:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) + 2)

            Case 5:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) + 3)
                      
        End Select

    Next

    DeTransform = sb
End Function

Function Transform(strPWD)
    Dim i
    i = 0
    Dim sb 
	sb = ""
    For i = 0 To Len(strPWD) - 1

        Select Case i Mod 6

            Case 0:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 8)

            Case 1:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 3)
     
            Case 2:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 9)

            Case 3:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 6)
         
            Case 4:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 2)

            Case 5:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 3)
                      
        End Select

    Next

    Transform = sb
End Function
%>

