VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReadConfigFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ReadCfgFile(ByVal strFilePath As String, Optional ByVal bolConvertToOneLine As Boolean = False) As String
        '<EhHeader>
        On Error GoTo ReadCfgFile_Err
        '</EhHeader>

        Dim iFile As Integer
        Dim strTmp As String
        Dim SB As clsStringBuilder
100     Set SB = New clsStringBuilder

102     iFile = FreeFile()

104     Open strFilePath For Input As #iFile

106     Do While Not EOF(iFile)
    
108         Line Input #iFile, strTmp
110         SB.Append strTmp
        Loop

112     Close #iFile
        
114     If bolConvertToOneLine Then
        
116         ReadCfgFile = Replace(Replace(SB.ToString, vbCr, ""), vbLf, "")
        Else
118         ReadCfgFile = SB.ToString
        End If

        '<EhFooter>
        Exit Function

ReadCfgFile_Err:
        WriteLog "Err.Number:" & Err.Number & "Err.Description" & Err.Description & ":" & strFilePath & " may not exists in " & _
           "HotelList_Bot.ClsReadConfigFile.ReadCfgFile"
        '</EhFooter>
End Function
