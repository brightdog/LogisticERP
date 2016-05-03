Attribute VB_Name = "modSleep"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

'*************************************************************************
'**函 数 名： MySleep
'**输    入： DealyTime(Double) 需延时的时间
'**输    出： 无
'**功能描述： 经过改造的延时器无凝滞，突破原先TIMER控件65.5秒的限制
'**增加跨天直接跳出的功能,就是不知道DATE这个函数的执行效率如何了。。。。2014-04-16
'*************************************************************************
Public Sub MySleep(DealyTime As Single)

    Dim TimerCount As Single
    Dim lastTimer  As Single
    Dim StartDate  As Date
    StartDate = Date
    lastTimer = Timer
    TimerCount = lastTimer + DealyTime '增加N秒
    While TimerCount - Timer > 0
        
        If Date = StartDate Then
            DoEvents
            Sleep 10

            DoEvents
        Else
            
            WriteLog "等待跨天了。。。强制等待5秒后跳出循环!"
            Dim i As Long

            For i = 1 To 500
                Sleep 10
                DoEvents
            Next
            Exit Sub
        End If

    Wend
End Sub


