Attribute VB_Name = "modUpdateThirdPartExpressInfo"
Option Explicit

Public Function UpdateThirdPartExpressInfobyClient(ByRef arrExpressNO() As String) As String
    
    Dim strExpressNO As String
    
    Dim i As Integer

    For i = 0 To UBound(arrExpressNO)

        If arrExpressNO(i) <> "" Then
        
            strExpressNO = strExpressNO & arrExpressNO(i)
            
            If i < UBound(arrExpressNO) Then
                strExpressNO = strExpressNO & "|"
            End If
        
        Else
        
        End If

    Next

    If strExpressNO <> "" Then
        Call VBA.Shell(App.path & "\ExpressBot.exe " & strExpressNO, vbHide)
    End If

End Function
