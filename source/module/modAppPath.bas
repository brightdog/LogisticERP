Attribute VB_Name = "modAppPath"
Option Explicit

Public APP_CONFIG_PATH As String

Public Sub Init_Path()
    '在工程启动时，必须在SUB MAIN中得到运行，否则路径无法获得
    APP_CONFIG_PATH = App.path & "\config\"

End Sub
