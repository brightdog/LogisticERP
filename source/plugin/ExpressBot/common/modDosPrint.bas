Attribute VB_Name = "modDosPrint"
Option Explicit

'��ֻҪ���򵥵Ĺ��ܣ���������ipconfig�ܰѻ��Զ�ȡ������������ûЧ����Ҫô���Ǵ�ƪ��ƪ�ĵ���
'�������������ʵ�֣���Ҫ WScript.Shell ���(WSHom.ocx)��cmd.exe��֧�֣������������Ǿ�û����
'����  msgbox dosprint("ipconfig")
Public Function DosPrint(ByVal strCommand As String, _
                         Optional ByVal bolGetOutPut As Boolean = False, _
                         Optional WaitOnReturn As Boolean = False) As String
        '<EhHeader>
        On Error GoTo DosPrint_Err
        '</EhHeader>
        Dim objShell As IWshRuntimeLibrary.WshShell
100     Set objShell = New IWshRuntimeLibrary.WshShell
        '    Dim objWshScriptExec As IWshRuntimeLibrary.WshExec
        '    Dim objStdOut As TextStream
        '    Set objWshScriptExec = objShell.Exec("c:\windows\system32\cmd.exe /q /c " & strCommand & " >> " & App.Path & "\log\log1111.txt")
        
102     If bolGetOutPut Then
            Dim strResult As String
            'WriteLog "++++++" & "c:\windows\system32\cmd.exe /q /c " & strCommand & " > """ & App.Path & "\log\tmp.tmp"""
            
104         strResult = objShell.Run("c:\windows\system32\cmd.exe /q /c " & strCommand & " > """ & App.path & "\log\tmp.tmp""", 0, WaitOnReturn)

            '
            '    Set objStdOut = objWshScriptExec.StdOut
            '
            '    DosPrint = objStdOut.ReadAll
    
            Dim iFreeFileNum As Integer

106         iFreeFileNum = FreeFile()
    
            
    
108         Open App.path & "\log\tmp.tmp" For Input As #iFreeFileNum
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
        
            'objShell.Run "c:\windows\system32\cmd.exe /q /c " & strCommand, 0, WaitOnReturn
124         Shell strCommand, vbNormalNoFocus
        
        End If
        
126     Set objShell = Nothing
        
        '<EhFooter>
        Exit Function

DosPrint_Err:

        WriteLog Err.Description & vbCrLf & "in modDosPrint.DosPrint " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

