Attribute VB_Name = "modgetMappingCode_Core"
Option Explicit

Public Function getMappingCode_Core(ByVal TypeFileName As String, ByVal strCode As String) As String
'"ThdVender"
'ThdVender.Mapping

    Dim Fso As Scripting.FileSystemObject
    Set Fso = New Scripting.FileSystemObject
    
    Dim strMappingFileName As String
    
    strMappingFileName = App.path & "\Config\" & TypeFileName & ".Mapping"
    
    If Fso.FileExists(strMappingFileName) Then
    
        Dim strContent As String
        
        strContent = Fso.OpenTextFile(strMappingFileName, ForReading, False, TristateFalse).ReadAll
        
        Dim arrLine() As String
        
        arrLine = Split(strContent, vbCrLf)
        
        Dim dicMapping As Scripting.Dictionary
        
        Set dicMapping = New Scripting.Dictionary
        
        Dim i As Integer
        
        For i = 0 To UBound(arrLine)
            If Trim(arrLine(i)) <> "" Then
            
                Dim arrTmp() As String
                arrTmp = Split(arrLine(i), vbTab, 2, vbBinaryCompare)
                If UBound(arrTmp) = 1 Then
                
                    If Not dicMapping.Exists(arrTmp(0)) Then
                    
                        dicMapping.Add arrTmp(0), arrTmp(1)
                    
                    Else
                    
                    
                    End If
                
                End If
            
            End If
        Next
        
        getMappingCode_Core = CStr(dicMapping.Item(strCode))
    
    Else
        getMappingCode_Core = ""
    End If
    
    Set Fso = Nothing
    
    If getMappingCode_Core = "" Then
        getMappingCode_Core = strCode
    End If
    
End Function
