Attribute VB_Name = "modSleep"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

'*************************************************************************
'**�� �� ���� MySleep
'**��    �룺 DealyTime(Double) ����ʱ��ʱ��
'**��    ���� ��
'**���������� �����������ʱ�������ͣ�ͻ��ԭ��TIMER�ؼ�65.5�������
'**���ӿ���ֱ�������Ĺ���,���ǲ�֪��DATE���������ִ��Ч������ˡ�������2014-04-16
'*************************************************************************
Public Sub MySleep(DealyTime As Single)

    Dim TimerCount As Single
    Dim lastTimer  As Single
    Dim StartDate  As Date
    StartDate = Date
    lastTimer = Timer
    TimerCount = lastTimer + DealyTime '����N��
    While TimerCount - Timer > 0
        
        If Date = StartDate Then
            DoEvents
            Sleep 10

            DoEvents
        Else
            
            WriteLog "�ȴ������ˡ�����ǿ�Ƶȴ�5�������ѭ��!"
            Dim i As Long

            For i = 1 To 500
                Sleep 10
                DoEvents
            Next
            Exit Sub
        End If

    Wend
End Sub


