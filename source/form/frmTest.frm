VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "frmTest"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7140
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdtestCust_Detail 
      Caption         =   "Test cust_ detail"
      Height          =   435
      Left            =   3060
      TabIndex        =   13
      Top             =   1500
      Width           =   1755
   End
   Begin VB.TextBox txtCustDistrict 
      Height          =   315
      Left            =   3660
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3780
      Width           =   1635
   End
   Begin VB.TextBox txtCustCity 
      Height          =   315
      Left            =   1920
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3780
      Width           =   1635
   End
   Begin VB.TextBox txtCustProvince 
      Height          =   315
      Left            =   180
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3780
      Width           =   1635
   End
   Begin VB.TextBox txtIncremental 
      Height          =   315
      Left            =   5400
      TabIndex        =   9
      Top             =   480
      Width           =   1635
   End
   Begin VB.ListBox lstIncremental 
      Height          =   1860
      Left            =   5640
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdShowOrderDetail 
      Caption         =   "Show order detail"
      Height          =   660
      Left            =   540
      TabIndex        =   7
      Top             =   1380
      Width           =   1815
   End
   Begin VB.CommandButton cmdTestLogin 
      Caption         =   "TestLogin"
      Height          =   375
      Left            =   3540
      TabIndex        =   0
      Top             =   600
      Width           =   1395
   End
   Begin VB.CommandButton cmdTestNetSendPic 
      Caption         =   "TestNetSendPic"
      Height          =   345
      Left            =   1740
      TabIndex        =   6
      Top             =   600
      Width           =   1500
   End
   Begin VB.CommandButton cmdTestNetSend 
      Caption         =   "TestNetSend"
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   1500
   End
   Begin VB.CommandButton cmdShowMainForm 
      Caption         =   "Show main form"
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   60
      Width           =   1500
   End
   Begin VB.TextBox txtCalendar 
      Height          =   375
      Left            =   1845
      TabIndex        =   3
      Text            =   "Calendar"
      Top             =   30
      Width           =   1590
   End
   Begin VB.CommandButton cmdTestCalender 
      Caption         =   "Test calender"
      Height          =   375
      Left            =   3555
      TabIndex        =   2
      Top             =   -15
      Width           =   1395
   End
   Begin VB.CommandButton cmdTestfrmDialog 
      Caption         =   "TestfrmDialog"
      Height          =   375
      Left            =   660
      TabIndex        =   1
      Top             =   2760
      Width           =   1395
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mLastIncrementalControl As VB.Control
Private Sub cmdShowMainForm_Click()
    frmMain.Show
End Sub

Private Sub cmdShowOrderDetail_Click()
    frmOrder_Detail.Show
End Sub

'Private Sub cmdShowMainForm_Click()
'    MDIFrmMain.Show
'End Sub

Private Sub cmdTestCalender_Click()
    
    Set frmCalender.ValueReturnControl = Me.txtCalendar
    frmCalender.MinDate = "2014-08-14"
    frmCalender.MaxDate = "2014-09-16"

    If IsDate(Me.txtCalendar.Text) Then
        frmCalender.SelectDate = Me.txtCalendar.Text
    End If

    frmCalender.Top = Me.cmdTestCalender.Top + Me.Top + Me.cmdTestCalender.Height
    frmCalender.Left = Me.cmdTestCalender.Left + Me.Left + Me.cmdTestCalender.width
    frmCalender.Show vbModal
    
End Sub

Private Sub cmdtestCust_Detail_Click()
    Load frmCust_Detail
    Call frmCust_Detail.LoadDetail(2)
    frmCust_Detail.Show
End Sub

Private Sub cmdTestfrmDialog_Click()
    frmDialog.msgBotton = 3
    frmDialog.msgTitle = "test"
    frmDialog.msgTxt = "abc123" & vbCrLf & "abc123" & vbCrLf & "abc123" & vbCrLf & "abc123" & vbCrLf & "abc123"
    frmDialog.ShowMsg
    frmDialog.Show vbModal
End Sub

Private Sub cmdTestLogin_Click()
    frmLogin.Show
    frmLogin.txtUserName.Text = "admin"
    frmLogin.txtPassword.Text = "admin"
    
End Sub

Private Sub cmdTestNetSend_Click()
    Dim obj As New clsNetOperator
    Dim dicData As New Scripting.Dictionary
    Dim objEncncodeURI As New clsEncncodeURI
    
    dicData.Add "URL", "http://114.215.177.126:9527/test.asp"
    dicData.Add "PostData", "a=b&1=" & modGetRandomNum.GetRandomNum(10, 6, 888888888) & "&chs=" & objEncncodeURI.UTF8EncodeURI("你好！")
    dicData.Add "Referer", "http://114.215.177.126:9527"
    
    Set dicData = obj.SendData(dicData)
    
    Debug.Print dicData.Item("ReturnCode")
    Debug.Print dicData.Item("ReturnData")
    
    Set obj = Nothing
    Set dicData = Nothing
End Sub

Private Sub cmdTestNetSendPic_Click()
Dim myUpload As vbsFileUpload
Set myUpload = New vbsFileUpload
myUpload.c_strDestURL = "http://114.215.177.126:9527/upload.asp"      ' 必选
myUpload.c_strFileName = "d:\testUpload.jpg"                                   ' 必选
'myUpload.c_strFileName = App.path & "\resource\testUpload.jpg"                                   ' 必选
myUpload.c_strFieldName = "file1"                                        ' 必选
myUpload.c_strContentType = "image/jpeg"                               ' 可选
Call myUpload.vbsUpload
Debug.Print myUpload.c_strResponseText
Debug.Print myUpload.c_strErrMsg
Set myUpload = Nothing
End Sub

Private Sub lstIncremental_Click()
    If Me.lstIncremental.Text <> "" Then
    
        mLastIncrementalControl.Text = Me.lstIncremental.Text
        Me.lstIncremental.Visible = False
    End If
End Sub

Private Sub txtCustProvince_Change()

    If Me.txtCustProvince.Text <> "" Then
        Set mLastIncrementalControl = Me.txtCustProvince
        Me.lstIncremental.Visible = QuichSearchbyLocaldic(Me.txtCustProvince, lstIncremental, gdicLocation.Item("0"))
    Else
        Me.lstIncremental.Visible = False
    End If

End Sub

Private Sub txtCustProvince_KeyDown(KeyCode As Integer, Shift As Integer)

    If Me.lstIncremental.Visible Then

        Select Case True

            Case KeyCode = vbKeyDown

                If Me.lstIncremental.ListIndex < Me.lstIncremental.ListCount - 1 Then
                    Me.lstIncremental.ListIndex = Me.lstIncremental.ListIndex + 1
                End If

            Case KeyCode = vbKeyUp

                If Me.lstIncremental.ListIndex > 0 Then
                    Me.lstIncremental.ListIndex = Me.lstIncremental.ListIndex - 1
                End If

            Case KeyCode = vbKeyReturn

                If Me.lstIncremental.ListIndex >= 0 Then
                    Me.txtCustProvince.Text = Me.lstIncremental.Text
                    Me.lstIncremental.Visible = False
                End If

        End Select

    End If

End Sub
Private Sub txtIncremental_Change()

    If Me.txtIncremental.Text <> "" Then
        Dim dicIncremental As Scripting.Dictionary
        lstIncremental.Clear
        Set dicIncremental = IncrementalSearch("tblOrder", "SenderName", "like", Me.txtIncremental.Text)
        Dim i As Integer

        If dicIncremental.Item("Rst").Count > 0 Then

            For i = 1 To dicIncremental.Item("Rst").Count
        
                lstIncremental.AddItem CStr(dicIncremental.Item("Rst").Item(i).Item(1))
    
            Next

            lstIncremental.Move Me.txtIncremental.Left, Me.txtIncremental.Top + Me.txtIncremental.Height, Me.txtIncremental.width, 1000
            Me.lstIncremental.Selected(0) = True
            Me.lstIncremental.Visible = True
        End If

    Else
        Me.lstIncremental.Visible = False
    End If

End Sub

Private Sub txtIncremental_KeyDown(KeyCode As Integer, Shift As Integer)

    If Me.lstIncremental.Visible Then

        Select Case True

            Case KeyCode = vbKeyDown

                If Me.lstIncremental.ListIndex < Me.lstIncremental.ListCount - 1 Then
                    Me.lstIncremental.ListIndex = Me.lstIncremental.ListIndex + 1
                End If

            Case KeyCode = vbKeyUp

                If Me.lstIncremental.ListIndex > 0 Then
                    Me.lstIncremental.ListIndex = Me.lstIncremental.ListIndex - 1
                End If

            Case KeyCode = vbKeyReturn

                If Me.lstIncremental.ListIndex >= 0 Then
                    Me.txtIncremental.Text = Me.lstIncremental.Text
                    Me.lstIncremental.Visible = False
                End If

        End Select

    End If

End Sub

