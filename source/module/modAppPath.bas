Attribute VB_Name = "modAppPath"
Option Explicit

Public APP_CONFIG_PATH As String

Public Sub Init_Path()
    '�ڹ�������ʱ��������SUB MAIN�еõ����У�����·���޷����
    APP_CONFIG_PATH = App.path & "\config\"

End Sub
