Attribute VB_Name = "modReadClipBoardData"
Option Explicit

Public Function GetRAWdataFromClipBrd() As String
    Dim strRawDataFromClipBrd As String
    strRawDataFromClipBrd = VB.Clipboard.GetText
    Debug.Print strRawDataFromClipBrd
    GetRAWdataFromClipBrd = strRawDataFromClipBrd
End Function

