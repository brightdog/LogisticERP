VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "ExpressBot"
   ClientHeight    =   4305
   ClientLeft      =   4215
   ClientTop       =   3480
   ClientWidth     =   8280
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   8280
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdEnd 
      Caption         =   "强关"
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Top             =   60
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5820
      Top             =   3840
   End
   Begin VB.TextBox txtCurrentName 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2940
      TabIndex        =   3
      Top             =   45
      Width           =   5325
   End
   Begin VB.ListBox lstLog 
      Appearance      =   0  'Flat
      Height          =   3450
      ItemData        =   "frmMain.frx":0CCA
      Left            =   0
      List            =   "frmMain.frx":0CD1
      TabIndex        =   1
      Top             =   420
      Width           =   8205
   End
   Begin VB.CommandButton cmdStop 
      Appearance      =   0  'Flat
      Caption         =   "停 止"
      Enabled         =   0   'False
      Height          =   345
      Left            =   1020
      TabIndex        =   4
      Top             =   45
      Width           =   1155
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "开 始"
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   915
   End
   Begin VB.Label lblVer 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.0.0.0"
      Height          =   255
      Left            =   6660
      TabIndex        =   2
      Top             =   4020
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Write By 老吴

Option Explicit

Dim mbolStop As Boolean



Private Sub cmdEnd_Click()
    End
End Sub

Private Sub cmdStart_Click()
        '<EhHeader>
        On Error GoTo cmdStart_Click_Err

        '</EhHeader>
        
100     gbolisGetDataFromWeb = False
        
        'Me.txtCurrentName.SetFocus
102     Me.cmdStop.Enabled = True

104     Me.cmdStart.Enabled = False
        
        Dim Rst As ADODB.Recordset
106     Set Rst = New ADODB.Recordset
            
        Dim obj As Object
108     Me.txtCurrentName.Text = ""
        
110     gstrLogDateTime = Format(Now(), "YYYY-MM-DD_HH-MM-SS")

        Dim strCommand   As String
        Dim arrCommand() As String
        Dim strCommandLine As String
112     strCommandLine = VBA.Command

        
114     Me.txtCurrentName.Text = VBA.Command
116     Me.Caption = Command
118     gstrSqlFileName = ""
        Dim FSo As Scripting.FileSystemObject
        
        Dim i As Integer
        Dim arrExpressNO() As String
120     arrExpressNO = Split(strCommandLine & "|", "|", -1, vbBinaryCompare)
        
122     For i = 0 To UBound(arrExpressNO)

124         If arrExpressNO(i) <> "" Then
126             arrCommand = Split(arrExpressNO(i), "_", 2, vbBinaryCompare)
        
128             If UBound(arrCommand) < 1 Then
            
130                 WriteLog "Command ERR!!"
132                 End
                End If
        
134             gstrSite = arrCommand(0)
136             gstrExpressNO = arrCommand(1)

138             Select Case True
        
                    Case gstrSite = "ZTO"
                
140                     gstrLogFileName = "Log_" & strCommandLine & "_" & gstrLogDateTime & ".txt"
142                     gstrSqlFileName = "Result_" & strCommandLine & ".txt"
144                     Set obj = New clsZTO
                
146                     Call obj.GetInfo(gstrExpressNO)
                        'strFileName = arrCommand(1)

148                 Case gstrSite = "STO"
                
150                     gstrLogFileName = "Log_" & strCommandLine & "_" & gstrLogDateTime & ".txt"
152                     gstrSqlFileName = "Result_" & strCommandLine & ".txt"
154                     Set obj = New clsSTO
                
156                     Call obj.GetInfo(gstrExpressNO)

158                 Case gstrSite = "UPS"
                
160                     gstrLogFileName = "Log_" & strCommandLine & "_" & gstrLogDateTime & ".txt"
162                     gstrSqlFileName = "Result_" & strCommandLine & ".txt"
164                     Set obj = New clsUPS
                
166                     Call obj.GetInfo(gstrExpressNO)

168                 Case Else
170                     WriteLog "参数非法！" & strCommandLine
                End Select

            End If

        Next

        '436     If strListFileName <> "" And gbolisGetDataFromWeb Then
        '438         If gstrSqlFileName <> "" Then
        '440             Call DosPrint("""C:\Program Files\7-Zip\7z.exe"" a D:\wwwroot\bibot\filelist\sqlfile\SQL_" & VBA.Command$ & "_" & gstrLogDateTime & ".7z -ppassword -mhe " & App.path & "\SQL\SQL_" & VBA.Command$ & "_" & gstrLogDateTime & ".sql")
        '            Else
        '442             Call DosPrint("""C:\Program Files\7-Zip\7z.exe"" a D:\wwwroot\bibot\filelist\sqlfile\SQL_" & strFileName & "_" & gstrLogDateTime & ".7z -ppassword -mhe " & App.path & "\SQL\SQL_" & strFileName & "_" & gstrLogDateTime & ".sql")
        '            End If
        '
        '        End If
        
        '136     gstrLogFileName = "Log_Taobao_" & Format(Date, "YYYY-MM-DD") & ".txt"
        '138     Set obj = New clsTaobao
        '140     Me.txtCurrentName.Text = "Taobao"
        '142     obj.GetInfo
172     Set FSo = Nothing
174     gstrLogFileName = ""
        
176     Me.txtCurrentName.Text = ""
178     Me.cmdStop.Enabled = False
180     Me.cmdStart.Enabled = True

182     End
        
        'MsgBox "DONE"
        
        '<EhFooter>
        Exit Sub

cmdStart_Click_Err:
        WriteLog Err.Description & vbCrLf & "in HotelPrice_Bot_ADSL.frmMain.cmdStart_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub



Private Sub cmdStop_Click()
    WriteLog "%%END%%"
    gbolExitJob = True
End Sub

Private Sub Form_Load()
    
    'Me.Show

    lblVer.Caption = App.Major & "." & App.Minor & "." & App.Revision

    Call cmdStart_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.lstLog.width = Me.width - Me.lstLog.Left * 2 - 100
    Me.lstLog.Height = Me.Height - Me.cmdStart.Top * 2 - Me.cmdStart.Height - 300 '- me.lblVer.Height
End Sub

Private Sub Timer1_Timer()

    If Format(Now, "hh:mm:ss") = "23:50:00" Then
        WriteLog "%%跨日的任务，不执行%%"
        gbolExitJob = True
    End If
    
    If CheckExitAllJob() Then
        WriteLog "%%检测到退出当前任务命令！%%"
        gbolExitJob = True
    End If
    

End Sub

Private Function CheckExitAllJob() As Boolean
    CheckExitAllJob = False
    Dim FSo As Scripting.FileSystemObject
    Set FSo = New Scripting.FileSystemObject

    If FSo.FileExists("D:\ServiceApp\GlobalADSLControl\EXITALL") Then
        CheckExitAllJob = True
    End If

    Set FSo = Nothing

End Function
