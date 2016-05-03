Attribute VB_Name = "modCalcTime"
Option Explicit

Public Function GuessFinishTime(ByVal PassedNum As Long, ByVal TotalNum As Long, ByVal AVGTime As String, Optional ByVal NeedFormat As Boolean = False) As String
        '返回预计完成需要多少秒时间。
        '<EhHeader>
        On Error GoTo GuessFinishTime_Err
        '</EhHeader>
    
        Dim strResult As String
    
100     strResult = CStr((TotalNum - PassedNum) * AVGTime)
        'strResult = CStr(TotalNum / (PassedNum / DateDiff("s", StartTime, VBA.Now())))
        If NeedFormat Then
        
            Select Case True
                
                Case strResult <= 0
                
                    
            
                Case strResult > 0 And strResult < 120
                
                
                Case strResult < 3600
                    strResult = Format(strResult / 60, "0.0") & " min "
                
                Case Else
                
                    Dim strHour As String
                    Dim strMin As String
                    
                    strHour = CInt(strResult / 3600)
                    strMin = CInt(strResult Mod 3600 / 60)
                    
                    strResult = strHour & " h " & strMin & " min "
            
            End Select
        
        End If
        


102     GuessFinishTime = strResult
        '<EhFooter>
        Exit Function

GuessFinishTime_Err:
        WriteLog Err.Description & vbCrLf & _
               "in HotelPrice_Bot_ADSL.modCalcTime.GuessFinishTime " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
