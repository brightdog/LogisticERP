Attribute VB_Name = "modCustERR"
Option Explicit


Public gERR As Boolean
Public gERRDESC As String

Public Sub ERRClear()
    gERR = False
    gERRDESC = ""
End Sub
