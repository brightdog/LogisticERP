VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form frmTransfer_Detail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "承运人详情"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "frmTransfer_Detail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8070
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cboTransferProvince 
      Height          =   300
      Left            =   1380
      TabIndex        =   3
      Text            =   "cboTransferProvince"
      Top             =   1860
      Width           =   1275
   End
   Begin VB.TextBox txtCreateEmp 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Create emp"
      Top             =   120
      Width           =   1935
   End
   Begin TDBDate6Ctl.TDBDate txtCreateDT 
      Height          =   315
      Left            =   4320
      TabIndex        =   10
      Top             =   120
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      Calendar        =   "frmTransfer_Detail.frx":000C
      Caption         =   "frmTransfer_Detail.frx":0107
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmTransfer_Detail.frx":016A
      Keys            =   "frmTransfer_Detail.frx":0188
      Spin            =   "frmTransfer_Detail.frx":01E6
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy-mm-dd"
      EditMode        =   3
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   1
      ForeColor       =   -2147483640
      Format          =   "yyyy-mm-dd"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   41640
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "2014-08-12"
      ValidateMode    =   0
      ValueVT         =   2010185735
      Value           =   41863
      CenturyMode     =   0
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   435
      Left            =   5100
      TabIndex        =   8
      Top             =   4620
      Width           =   1275
   End
   Begin VB.TextBox txtRemark 
      Height          =   555
      Left            =   1380
      TabIndex        =   5
      Text            =   "Remark"
      Top             =   3060
      Width           =   4035
   End
   Begin VB.TextBox txtTransferID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Transfer iD"
      Top             =   120
      Width           =   2535
   End
   Begin VB.ComboBox cboTransferType 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmTransfer_Detail.frx":020E
      Left            =   1380
      List            =   "frmTransfer_Detail.frx":0210
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3780
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveDetail 
      Caption         =   "保存"
      Height          =   435
      Left            =   1680
      TabIndex        =   7
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox txtTransferAddress 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1380
      TabIndex        =   4
      Text            =   "Transfer address"
      Top             =   2220
      Width           =   4575
   End
   Begin VB.TextBox txtTransferPhone 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      Text            =   "Transfer phone"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtTransferContactor 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      TabIndex        =   2
      Text            =   "Transfer contactor"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtTransferName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Text            =   "Transfer name"
      Top             =   540
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "承运人ID"
      Height          =   195
      Left            =   60
      TabIndex        =   18
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "承运人名称"
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "承运人电话"
      Height          =   195
      Left            =   60
      TabIndex        =   16
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "承运人联系人"
      Height          =   195
      Left            =   60
      TabIndex        =   15
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "承运人地址"
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "备注信息"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   3060
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "承运人类型"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "frmTransfer_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dicLayout As Scripting.Dictionary
Dim mbolLayout As Boolean
Public mID As Long
Private bolisNew As Boolean

Public Function LoadDetail(ByVal ID As String) As String
    bolisNew = False
    mID = ID
    Dim strURL, strPostData As String
    strPostData = "data={""Type"":""TransferDetail"",""Fields"":[""TransferID""],""Values"":[""" & ID & """]}"
    strURL = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strURL, strPostData)
    Call FillFormTextBox(Me, JSON.Parse(strResult))
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSaveDetail_Click()

    If bolisNew Then
        Me.txtCreateDT.Text = Format(VBA.Now(), "yyyy-mm-dd hh:mm:ss")
    End If

    Dim ctl As VB.Control
    Dim SBField As clsStringBuilder
    Set SBField = New clsStringBuilder
    Dim SBValue As clsStringBuilder
    Set SBValue = New clsStringBuilder

    For Each ctl In Me.Controls

        'Debug.Print "sAVe:" & ctl.name & ":" & ctl.Value
        If isCtlLinkedDB(ctl) Then
            
            SBField.Append """" & Right(ctl.name, Len(ctl.name) - 3) & ""","

            If TypeName(ctl) = "CheckBox" Or TypeName(ctl) = "TDBDate" Then
                SBValue.Append """" & MakeQueryValue(Right(ctl.name, Len(ctl.name) - 3), ctl.Value & "", False) & ""","
            Else
                Debug.Print ctl.name
                SBValue.Append """" & MakeQueryValue(Right(ctl.name, Len(ctl.name) - 3), ctl.Text, False) & ""","
            End If
        End If
    
    Next
    
    Dim strPostData As String
    Dim strURL As String
    Dim strFields As String
    Dim strValues As String
    strFields = SBField.toString
    strValues = SBValue.toString

    strPostData = "{""Type"":""OrderDetail"",""Fields"":[" & Left(strFields, Len(strFields) - 1) & "],""Values"":[" & Left(strValues, Len(strValues) - 1) & "]}"
    strURL = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strURL, strPostData)
    Debug.Print strResult
Dim dicResult As Scripting.Dictionary
    
    Set dicResult = JSON.Parse(strResult)
    
    If dicResult.Item("STATE") <> "ERR" Then
    'If Left(strResult, 1) <> "{" Then '流氓判断法，先把信息提示出来再说了。
        MsgBox "保存成功"
        Unload Me
    Else
        MsgBox "保存失败，请检查字段信息"
        
    End If

End Sub

Private Sub cmdEditLayout_Click()
    Dim ctl As VB.Control
    Set dicLayout = New Scripting.Dictionary

    For Each ctl In Me.Controls
    
        If TypeName(ctl) = "TextBox" Then
            
            ctl.MousePointer = 15
            ctl.Locked = True
            Dim dicSinglePos As Scripting.Dictionary
            Set dicSinglePos = New Scripting.Dictionary
            dicSinglePos.Add "Top", ctl.Top
            dicSinglePos.Add "Left", ctl.Left
            dicSinglePos.Add "X", 0
            dicSinglePos.Add "Y", 0
            dicSinglePos.Add "OffsetX", 0
            dicSinglePos.Add "OffsetY", 0
            
            dicLayout.Add ctl.name, dicSinglePos
            Set dicSinglePos = Nothing
        End If
    
    Next

    mbolLayout = True
    
End Sub

Private Sub cmdSaveLayout_Click()
    Dim ctl As VB.Control
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder

    For Each ctl In Me.Controls
    
        If TypeName(ctl) = "TextBox" Then
            SB.Append """" & ctl.name & """:{"
            SB.Append """Top"":" & """" & ctl.Top & ""","
            SB.Append """Left"":" & """" & ctl.Left & ""","
            SB.Append """Height"":" & """" & ctl.Height & ""","
            SB.Append """Width"":" & """" & ctl.width & """},"
            ctl.MousePointer = 0
            ctl.Locked = True
        End If
    
    Next
    
    Dim strContent As String
    strContent = SB.toString
    strContent = Left(strContent, Len(strContent) - 1)
    
    Dim Fso As Scripting.FileSystemObject
    Set Fso = New Scripting.FileSystemObject
    Dim oText As Scripting.TextStream
    Set oText = Fso.OpenTextFile(APP_CONFIG_PATH & Me.name & ".Layout", ForWriting, True, TristateFalse)
    
    oText.Write "{" & strContent & "}"

    oText.Close
    Set SB = Nothing
    Set Fso = Nothing
    mbolLayout = False
End Sub

Private Sub Form_Load()
    bolisNew = True
    Call InitLayout(Me)
    Call InitTextBox(Me)
    Me.cboTransferType.AddItem "我司"
    Me.cboTransferType.AddItem "三方"
    
    Call FillCboWithSampleDic(Me.cboTransferProvince, gdicLocation.Item("0"))
    Me.txtTransferID.Text = ""
    
End Sub

'======================================================================================================================
'======================================================================================================================
'======================================================================================================================
'======================================================================================================================
'======================================================================================================================
'
'Private Function MoveButton(ByRef btn As VB.TextBox, ByVal x As Single, ByVal Y As Single)
'
'    If dicLayout.Exists(btn.name) Then
'
'        btn.Move btn.Left + x - dicLayout.Item(btn.name).Item("OffsetX"), btn.Top + Y - dicLayout.Item(btn.name).Item("OffsetY")
'
'        DoEvents
'        dicLayout.Item(btn.name).Item("Top") = btn.Top
'        dicLayout.Item(btn.name).Item("Left") = btn.Left
'        dicLayout.Item(btn.name).Item("X") = 0
'        dicLayout.Item(btn.name).Item("Y") = 0
'        dicLayout.Item("txtReceiverName").Item("OffsetX") = 0
'        dicLayout.Item("txtReceiverName").Item("OffsetY") = 0
'    End If
'
'End Function
'
'
'
'
'Private Sub txtReceiverName_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'
'    If mbolLayout Then
'        If dicLayout.Exists("txtReceiverName") Then
'            Debug.Print x & ":::" & Y
'            dicLayout.Item("txtReceiverName").Item("X") = x
'            dicLayout.Item("txtReceiverName").Item("Y") = Y
'            dicLayout.Item("txtReceiverName").Item("OffsetX") = x - dicLayout.Item("txtReceiverName").Item("Left")
'            dicLayout.Item("txtReceiverName").Item("OffsetY") = Y - dicLayout.Item("txtReceiverName").Item("Top")
'
'        End If
'    End If
'
'End Sub
'
'Private Sub txtReceiverName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If mbolLayout Then
'    If Button = vbLeftButton Then
'        If dicLayout.Exists("txtReceiverName") Then
'            Debug.Print x & ":" & Y
'            Call MoveButton(Me.txtReceiverName, x - CSng(dicLayout.Item("txtReceiverName").Item("X")), Y - CSng(dicLayout.Item("txtReceiverName").Item("Y")))
'            'DoEvents
'        End If
'    End If
'End If
'End Sub
