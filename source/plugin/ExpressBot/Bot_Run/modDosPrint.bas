Attribute VB_Name = "modDosPrint"
Option Explicit

'��ֻҪ���򵥵Ĺ��ܣ���������ipconfig�ܰѻ��Զ�ȡ������������ûЧ����Ҫô���Ǵ�ƪ��ƪ�ĵ���
'�������������ʵ�֣���Ҫ WScript.Shell ���(WSHom.ocx)��cmd.exe��֧�֣������������Ǿ�û����
'����  msgbox dosprint("ipconfig")
Public Function DosPrint(ByVal strCommand As String, Optional ByVal bolGetOutPut As Boolean = True) As String
        '<EhHeader>
        On Error GoTo DosPrint_Err
        '</EhHeader>
        Dim objShell As IWshRuntimeLibrary.WshShell
100     Set objShell = New IWshRuntimeLibrary.WshShell
        '    Dim objWshScriptExec As IWshRuntimeLibrary.WshExec
        '    Dim objStdOut As TextStream
        '    Set objWshScriptExec = objShell.Exec("c:\windows\system32\cmd.exe /q /c " & strCommand & " >> " & App.Path & "\log\log1111.txt")
        
        If bolGetOutPut Then
            
            'WriteLog "++++++" & "c:\windows\system32\cmd.exe /q /c " & strCommand & " > """ & App.Path & "\log\tmp.tmp"""
            
102         objShell.Run "c:\windows\system32\cmd.exe /q /c " & strCommand & " > """ & App.Path & "\log\tmp.tmp""", 0, True
104
            '
            '    Set objStdOut = objWshScriptExec.StdOut
            '
            '    DosPrint = objStdOut.ReadAll
    
            Dim iFreeFileNum As Integer

106         iFreeFileNum = FreeFile()
    
            Dim strResult As String
    
108         Open App.Path & "\log\tmp.tmp" For Input As #iFreeFileNum
            Dim strTmp As String

110         Do While Not EOF(iFreeFileNum)
        
112             Line Input #iFreeFileNum, strTmp

114             If strResult = "" Then

116                 strResult = strTmp

                Else

118                 strResult = strResult & vbCrLf & strTmp

                End If

            Loop

120         Close #iFreeFileNum
    
122         DosPrint = strResult
        
        Else
        
            objShell.Run "c:\windows\system32\cmd.exe /q /c " & strCommand, 0, True
        
        End If
        
        Set objShell = Nothing
        
        '<EhFooter>
        Exit Function

DosPrint_Err:

        WriteLog Err.Description & vbCrLf & _
           "in WFUpload_Client.modDosPrint.DosPrint " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

