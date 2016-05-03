Attribute VB_Name = "modFormConfig"
Option Explicit

Public Function SaveConfig(ByRef frm As VB.Form, Optional ByVal strConfigFileName As String = "") As Boolean
        '<EhHeader>
        On Error GoTo SaveConfig_Err
        '</EhHeader>

        Dim ctl As Control
        Dim strConfig As String
        
        Dim clsEcypt As clsEncrypt
100     Set clsEcypt = New clsEncrypt
        
102     For Each ctl In frm.Controls
        
104         If TypeName(ctl) = "TextBox" Then
106             If ctl.name <> "txtLog" Then
108                 strConfig = strConfig & clsEcypt.Encode(ctl.name & "|-|" & Replace(ctl.Text, vbCrLf, "|--|")) & vbCrLf
                End If
            End If
        
        Next
        
110     If strConfig <> "" Then

            Dim iFile As Integer

112         iFile = FreeFile()
114         If strConfigFileName <> "" Then
            
116             Open App.path & "\config\" & clsEcypt.Encode(strConfigFileName) For Output As #iFile
            Else
118             Open App.path & "\config\" & clsEcypt.Encode(frm.name) For Output As #iFile
            End If

120         Print #iFile, Left(strConfig, Len(strConfig) - 2)
122         Close #iFile
                    
124         SaveConfig = True
                    
        End If
        Set clsEcypt = Nothing
        '<EhFooter>
        Exit Function

SaveConfig_Err:
        WriteLog Err.Description & vbCrLf & _
           "in LogisticERP.modConfig.SaveConfig " & _
           "at line " & Erl
        SaveConfig = False
        '</EhFooter>
End Function

Public Function ReadConfig(ByRef frm As VB.Form, Optional ByVal strConfigFileName As String = "") As Boolean

        '<EhHeader>
        On Error GoTo ReadConfig_Err
        '</EhHeader>

        Dim iFile As Integer
        Dim clsEcypt As clsEncrypt
110     Set clsEcypt = New clsEncrypt
        
100     iFile = FreeFile()
102     If strConfigFileName <> "" Then
            
104         Open App.path & "\config\" & clsEcypt.Encode(strConfigFileName) For Input As #iFile
        Else
106         Open App.path & "\config\" & clsEcypt.Encode(frm.name) For Input As #iFile
        End If
    
        'MsgBox App.Path
    
        Dim I As Integer
108     I = 1
    

112     Do While Not EOF(1)
            Dim strTmp As String
            Dim arr() As String
114         Line Input #iFile, strTmp
116         arr = Split(clsEcypt.Decode(strTmp), "|-|", 2, vbBinaryCompare)

118         If UBound(arr) = 1 Then
        
120             CallByName frm, arr(0), VbLet, restoreCRLF(arr(1))
        
            Else
        
            End If
        
122         I = I + 1
        Loop

124     Close #iFile
        Set clsEcypt = Nothing
126     ReadConfig = True
        '<EhFooter>
        Exit Function

ReadConfig_Err:
        WriteLog Err.Description & vbCrLf & _
           "in LogisticERP.modConfig.ReadConfig " & _
           "at line " & Erl
        ReadConfig = False
        
        '</EhFooter>
End Function

