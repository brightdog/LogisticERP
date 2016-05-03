VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Bot_RunAutoPlus"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   7425
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrUpload 
      Interval        =   2000
      Left            =   4260
      Top             =   180
   End
   Begin VB.Timer Tmr 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6600
      Top             =   1200
   End
   Begin VB.VScrollBar vsThreadNum 
      Height          =   315
      Left            =   2520
      Max             =   49
      TabIndex        =   5
      Top             =   60
      Value           =   40
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtThreadNum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "1"
      Top             =   60
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstExpressNO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   840
      TabIndex        =   3
      Top             =   540
      Width           =   2235
   End
   Begin VB.CommandButton cmdManualRun 
      Caption         =   "Run"
      Height          =   315
      Left            =   5880
      TabIndex        =   1
      Top             =   60
      Width           =   645
   End
   Begin VB.TextBox txtExpressNO 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "06:00:00|16:00:00"
      Top             =   420
      Width           =   6495
   End
   Begin VB.Label lblState 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bolStop As Boolean
Dim gdicTask As Scripting.Dictionary
Dim gstrTSKFile As String

Private Function getLocalExpressNO() As Scripting.Dictionary

        Dim Fso As Scripting.FileSystemObject
100     Set Fso = New Scripting.FileSystemObject

        Dim dicResult As Scripting.Dictionary
    
102     Set dicResult = New Scripting.Dictionary

104     If Fso.FileExists(gstrTSKFile) Then
    
            Dim strTmp As String
            Dim TS As Scripting.TextStream
        
106         Set TS = Fso.OpenTextFile(gstrTSKFile, ForReading, True, TristateFalse)
            Dim i As Integer
108         i = 0

110         Do While Not TS.AtEndOfStream
        
112             strTmp = TS.ReadLine
                Dim arr() As String
114             arr = Split(strTmp, vbTab, 2, vbBinaryCompare)
            
116             If UBound(arr) = 1 Then

118                 dicResult.Add i, arr
120                 i = i + 1
                Else
            
                End If
        
            Loop
    
        Else
    
        End If

122     Set getLocalExpressNO = dicResult
124     Set Fso = Nothing

End Function

Private Sub cmdManualRun_Click()
        On Error Resume Next

100     If Me.cmdManualRun.Caption = "Run" Then
102         Me.cmdManualRun.Caption = "Stop"
104         bolStop = False
            Dim i As Long
106         i = 0
            Dim dicTask As Scripting.Dictionary
108         Set dicTask = getLocalExpressNO()
        
            Dim dicExpressNO As Scripting.Dictionary
110         Set dicExpressNO = New Scripting.Dictionary

112         Me.lstExpressNO.Clear
            Dim v As Variant
        
114         For Each v In dicTask.keys
116             Me.lstExpressNO.AddItem dicTask.Item(v)(0)
            Next
        
118         Me.txtExpressNO.Visible = False
120         Me.lstExpressNO.Visible = True
122         i = 0
        
            Dim Reg As VBScript_RegExp_55.RegExp
124         Set Reg = New VBScript_RegExp_55.RegExp
        
126         Reg.Global = True
128         Reg.MultiLine = False
130         Reg.IgnoreCase = True
        
132         Reg.Pattern = "(\d+)"
        
            Dim Mc As VBScript_RegExp_55.MatchCollection
        
134         For Each v In dicTask.keys
            
136             Set Mc = Reg.Execute(dicTask.Item(v)(0))
            
138             If Mc.Count > 0 Then
                    Dim strSeed As String
                
140                 strSeed = getLongestNum(Mc)
142                 dicExpressNO.Add Replace(dicTask.Item(v)(0), strSeed, "|-SEED-|", 1, 1, vbBinaryCompare) & "%" & strSeed, Array(0, dicTask.Item(v)(1))
            
                End If
            
            Next
        
            Dim strAddNum As String
144         strAddNum = 0
            Dim strAddCount As String
146         strAddCount = 1

148         Do While Not bolStop
            
150             If modProc.GetProcessCountbyName("ExpressBot.exe") < CInt(Me.txtThreadNum.Text) Then
                    '=========================================
                    '根据列表中的单号，随机取一个拿出来用。
152                 Call VBA.Randomize

                    Dim pickUpNum As Integer

154                 pickUpNum = Int(VBA.Rnd() * dicTask.Count + 1)

                    Dim strExpressNO As String
156                 strExpressNO = getRandExpressNO(dicExpressNO, pickUpNum)
                    Dim arrTask() As String
158                 arrTask = Split(strExpressNO, vbTab, 2, vbBinaryCompare)
160                 strExpressNO = arrTask(0)
                    Dim ADDLimit As String
162                 ADDLimit = arrTask(1)
                    '=========================================
                    Dim arrExpressNO() As String
164                 arrExpressNO = Split(strExpressNO, "%", 2, vbBinaryCompare)

166                 If UBound(arrExpressNO) = 1 Then
                        Dim strConstPart As String
                
                        Dim strAddPart As String
                
168                     strConstPart = arrExpressNO(0)
170                     strAddPart = arrExpressNO(1)
                    
                        Dim strAddNewValue As String

                        '                    strAddNewValue = BigADD(strAddPart, 1)

172                     If CDbl(ADDLimit) < CDbl(strAddCount) Then
174                         strAddPart = ""
                        End If
                    
176                     If strAddPart <> "" Then
                            Dim strDynamicPart As String
                        
178                         strExpressNO = Replace(strConstPart, "|-SEED-|", strAddPart, 1, 1, vbBinaryCompare)
180                         strDynamicPart = BigADD(strAddPart, 1)
182                         dicExpressNO.Remove CStr(arrExpressNO(0) & "%" & strAddPart)
184                         dicExpressNO.Add strConstPart & "%" & strDynamicPart, Array(strAddCount, ADDLimit)
186                         Shell App.Path & "\ExpressBot.exe " & strExpressNO, vbMinimizedNoFocus
                            Dim dicWrite As Scripting.Dictionary
188                         Set dicWrite = New Scripting.Dictionary
190                         dicWrite.RemoveAll
192                         Debug.Print ADDLimit - strAddCount
194                         dicWrite.Add Replace(strConstPart, "|-SEED-|", strDynamicPart, 1, 1, vbBinaryCompare), ADDLimit - strAddCount
196                         Call WriteTaskInfo(dicWrite)
198                         strAddCount = strAddCount + 1
                        Else
                        
200                         dicExpressNO.Item(arrExpressNO(0)) = strConstPart & "@FINISH"
                            Dim Fso As Scripting.FileSystemObject
202                         Set Fso = New Scripting.FileSystemObject
                            
204                         Call Fso.DeleteFile(gstrTSKFile, True)
                            Me.cmdManualRun.Caption = "Run"
206                         Set Fso = Nothing
                            
208                         Tmr.Enabled = True
                            Exit Sub
                    
                        End If

210                     Call fillList(Me.lstExpressNO, dicExpressNO)
212                     i = i + 1
                
                        'strAddNum = BigADD(strAddNum, 1)
                    End If

                Else
            
                    Do
214                     MySleep 0.1
216                 Loop While modProc.GetProcessCountbyName("ExpressBot.exe") >= CInt(Me.txtThreadNum.Text)
            
                End If
            
218             Me.lblState.Caption = i
220             MySleep 0.3
222             Debug.Print "x"
            Loop
        
        Else
    
224         Me.cmdManualRun.Caption = "Run"
226         bolStop = True
    
        End If

228     Me.txtExpressNO.Visible = True
230     Me.lstExpressNO.Visible = False
        
End Sub

Private Sub fillList(ByRef lst As VB.ListBox, ByRef dicExpressNO As Scripting.Dictionary)
        On Error Resume Next
        Dim v As Variant
100     lst.Clear

102     For Each v In dicExpressNO.keys
104     Debug.Print dicExpressNO.Item(v)(1) - dicExpressNO.Item(v)(0)
106         lst.AddItem CStr(v) & ":" & dicExpressNO.Item(v)(1) - dicExpressNO.Item(v)(0)
            'Debug.Print dicExpressNO.Item(v)(1) - dicExpressNO.Item(v)(0)
        Next

End Sub

Private Function getRandExpressNO(ByVal dicExpressNO As Scripting.Dictionary, ByVal pickUpNum As Integer) As String
        On Error Resume Next

        Dim v As Variant
        Dim i As Integer
    
100     i = 1

102     For Each v In dicExpressNO.keys
        
104         If i = pickUpNum Then
        
106             getRandExpressNO = CStr(v) & vbTab & dicExpressNO.Item(v)(1)
                Exit For
            End If

108         i = i + 1
        Next

End Function

Private Function getLongestNum(ByRef Mc As VBScript_RegExp_55.MatchCollection) As String
        On Error Resume Next
        Dim m As VBScript_RegExp_55.Match
        Dim dblbigger As String
100     dblbigger = 0

102     For Each m In Mc

104         If Len(CStr(m)) > Len(dblbigger) Then
        
106             dblbigger = m
        
108         ElseIf Len(CStr(m)) = Len(dblbigger) Then

110             If CStr(m) > dblbigger Then
112                 dblbigger = m
                End If
            End If

        Next

114     getLongestNum = dblbigger
End Function

Private Sub Form_Load()
    
100     gstrTSKFile = App.Path & "\LastTask.tsk"
102     Call Form_Resize
104     Call CheckTask
    
End Sub
Private Sub Tmr_Timer()
100     Call CheckTask
End Sub

Private Sub CheckTask()
100     Me.Show
102     Set gdicTask = GetLocalTask

104     If gdicTask.Count < 1 Then
    
106         Set gdicTask = GetServerTask
        
108         If gdicTask.Count > 0 Then
110             Tmr.Enabled = False
112             Call WriteTaskInfo(gdicTask)
114             Call cmdManualRun_Click
        
            Else
        
116             Tmr.Enabled = True
        
            End If

        Else
    
118         Tmr.Enabled = False
            'Call WriteTaskInfo(gdicTask)
120         Call cmdManualRun_Click
    
        End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
100     End
End Sub

Private Sub Form_Resize()
    
        On Error Resume Next
    
100     Me.txtExpressNO.Height = Me.Height - Me.lblState.Height - 450
102     Me.txtExpressNO.width = Me.width - 230
104     Me.lstExpressNO.Visible = False
106     Me.lstExpressNO.Top = Me.txtExpressNO.Top
108     Me.lstExpressNO.Left = Me.txtExpressNO.Left
110     Me.lstExpressNO.width = Me.txtExpressNO.width
112     Me.lstExpressNO.Height = Me.txtExpressNO.Height

114     If Me.cmdManualRun.Caption = "Stop" Then
116         Me.lstExpressNO.Visible = True
        End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
100     End
End Sub


Private Sub tmrUpload_Timer()
    
100     Call modZipnUpload.ZipnUploadResultFile


End Sub





Private Sub vsThreadNum_Change()
100     Me.txtThreadNum.Text = 50 - Me.vsThreadNum.value
End Sub

Private Function GetLocalTask() As Scripting.Dictionary
    
        Dim Fso As Scripting.FileSystemObject
100     Set Fso = New Scripting.FileSystemObject
        Dim dicResult As Scripting.Dictionary
102     Set dicResult = New Scripting.Dictionary
    
104     If Fso.FileExists(gstrTSKFile) Then
    
            Dim tmp As String
106         tmp = Fso.OpenTextFile(gstrTSKFile, ForReading, True, TristateFalse).ReadAll
            Dim arrLine() As String
        
108         arrLine = Split(tmp, vbCrLf, -1, vbBinaryCompare)
        
110         If UBound(arrLine) > -1 Then
                Dim i As Integer

112             For i = 0 To UBound(arrLine)

114                 If arrLine(i) <> "" Then
                        Dim arr() As String
        
116                     arr = Split(arrLine(i), vbTab, 2, vbBinaryCompare)
        
118                     If UBound(arr) = 1 Then
        
120                         dicResult.Add arr(0), arr(1)
        
                        Else
        
                        End If
                    End If

                Next

            End If

        Else
    
        End If
    
122     Set GetLocalTask = dicResult
    
124     Set Fso = Nothing

End Function

Private Function GetServerTask() As Scripting.Dictionary

        Dim iWeb As clsXMLHTTPGetHtml
100     Set iWeb = New clsXMLHTTPGetHtml
        Dim dicResult As Scripting.Dictionary
102     Set dicResult = New Scripting.Dictionary
    
104     iWeb.CharSet = "UTF-8"
106     iWeb.Url = GetServerAddressFromFile() ' "http://218.17.224.215:8081/sfbi/querycompany?company=ups"
        Dim strResult As String
108     strResult = iWeb.StartGetHtml
    
110     If Left(strResult, 1) = "{" Then
    
112         Set dicResult = JSON.Parse(strResult)
114         If dicResult.Exists("waybillno") And dicResult.Exists("addcnt") Then
        
                Dim strNO As String
                Dim strADD As String
            
116             strNO = VBA.UCase(dicResult.Item("companyname")) & "_" & dicResult.Item("waybillno")
118             strADD = dicResult.Item("addcnt")
120             Set dicResult = New Scripting.Dictionary
122             dicResult.Add strNO, strADD
        
            Else
        
124             WriteLog "返回JSON格式不正确"
            End If
    
        Else
126         WriteLog "运单请求返回非法"
        End If
    
128     Set iWeb = Nothing
130     Set GetServerTask = dicResult
    
End Function

Private Function GetServerAddressFromFile() As String

        Dim Fso As Scripting.FileSystemObject
100     Set Fso = New Scripting.FileSystemObject
    
102     If Fso.FileExists(App.Path & "\server.config") Then
    
            Dim TS As Scripting.TextStream
        
104         Set TS = Fso.OpenTextFile(App.Path & "\server.config", ForReading, False, TristateFalse)

106         If Not TS.AtEndOfStream Then
        
108             GetServerAddressFromFile = TS.ReadLine
        
            Else
110             GetServerAddressFromFile = ""
            End If
    
        Else
112         GetServerAddressFromFile = ""
        End If
    
114     Set Fso = Nothing

End Function

Private Sub WriteTaskInfo(ByRef dic As Scripting.Dictionary)

        Dim Fso As Scripting.FileSystemObject
100     Set Fso = New Scripting.FileSystemObject

        Dim v As Variant
        Dim strResult As String
102     strResult = ""
104     For Each v In dic.keys
    
106         strResult = strResult & v & vbTab & dic.Item(v) & vbCrLf
    
        Next
    
108     Call Fso.OpenTextFile(gstrTSKFile, ForWriting, True, TristateFalse).Write(strResult)
    
110     Set Fso = Nothing

End Sub
