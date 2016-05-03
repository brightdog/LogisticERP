Attribute VB_Name = "modGetRandomNum"
Option Explicit


Public Function GetRandomNum(ByRef intMaxLength As Integer, Optional ByVal intMinLength As Integer = 0, Optional ByVal intMax As Double = 0) As String

    If intMinLength > intMaxLength Then
        GetRandomNum = "00000000"
        Exit Function
    End If
    
    
reGet:
    Dim i As Integer
    Dim strResult As String
    strResult = ""
    Dim NumCount As Integer
    If intMinLength > 0 Then
        NumCount = intMaxLength - intMinLength
        Randomize Second(Now)
        NumCount = Round(NumCount * Rnd) + intMinLength
    Else
    
        NumCount = intMaxLength
    End If
    For i = 1 To NumCount
    
    
        Randomize Second(Now)
'        Dim tmp As Single
'        tmp = Rnd
        'Debug.Print Rnd & ":" & CInt(9# * Rnd)
reGetFirstNum:
        strResult = strResult & CStr(CInt(9# * Rnd))
    
        If i = 1 And strResult = "0" Then
            strResult = ""
            GoTo reGetFirstNum
        End If
    
    Next
    
    If intMax > 0 And CDbl(strResult) > intMax Then
        GoTo reGet
    End If

    GetRandomNum = Format(strResult, 0)

End Function
