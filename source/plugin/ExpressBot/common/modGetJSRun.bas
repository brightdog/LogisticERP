Attribute VB_Name = "modGetJSRun"
Option Explicit


Public Function GetJSRun(ByVal JSPath As String, Optional ByVal UA As String = "", Optional ByVal bolShowForm As Boolean = False) As String
        '<EhHeader>
        On Error GoTo GetJSRun_Err
        '</EhHeader>

100     If UA = "" Then
102         UA = "Mozilla/5.0 (Windows NT 5.1; rv:35.0) Gecko/20100101 Firefox/35.0"
        End If

        Dim objNewWB As frmWB
104     Set objNewWB = New frmWB
                    
106     Load objNewWB
108     objNewWB.WB.AddressBar = False
110     objNewWB.WB.MenuBar = False
112     objNewWB.WB.RegisterAsDropTarget = False
114     objNewWB.WB.TheaterMode = False
116     objNewWB.WB.Silent = True
118     objNewWB.WB.Resizable = False

120     If bolShowForm Then
122         objNewWB.Show
        End If
                    
124     WriteLog "StartDecodeJS:" & JSPath
                    
126     objNewWB.WB.Navigate JSPath, "_self", Nothing, , "User-Agent: " & UA
        Dim iWBWait As Integer
128     iWBWait = 0
                    
130     Do While objNewWB.WB.readyState <> READYSTATE_COMPLETE
132         WriteLog "WB.readyState=" & frmWB.WB.readyState
134         iWBWait = iWBWait + 1
                    
136         MySleep 0.01
                    
138         If iWBWait > 200 Then
140             WriteLog "iWBWait > 200"
142             GetJSRun = ""
                Exit Function
            End If
        Loop
                    
144     Do While Right(objNewWB.WB.Document.getElementById("txtResult").value, 5) <> "<END>"
            'WriteLog "Do While:" & iWBWait
146         iWBWait = iWBWait + 1
                    
148         MySleep 0.01
                    
150         If iWBWait > 200 Then
152             WriteLog "iWBWait > 200"
154             GetJSRun = ""
                Exit Function
            End If
                    
        Loop

156     GetJSRun = objNewWB.WB.Document.getElementById("txtResult").value
158     GetJSRun = Replace(Replace(GetJSRun, vbCr, ""), vbLf, "")
160     Set objNewWB = Nothing

        '<EhFooter>
        Exit Function

GetJSRun_Err:
        WriteLog "GetJSRun @ERL:" & Erl & ":" & Err.Description
        '</EhFooter>
End Function
