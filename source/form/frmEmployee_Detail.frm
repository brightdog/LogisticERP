VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form frmEmployee_Detail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "员工信息详情"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   Icon            =   "frmEmployee_Detail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   9090
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtPWD 
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
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   19
      Text            =   "PWD"
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox txtEngName 
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
      Left            =   1680
      TabIndex        =   17
      Text            =   "EngName"
      Top             =   540
      Width           =   3495
   End
   Begin VB.ComboBox txtEmployeeProvince 
      Height          =   300
      Left            =   1680
      TabIndex        =   16
      Top             =   2220
      Width           =   1575
   End
   Begin VB.TextBox txtEmployeeAddress 
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
      Height          =   435
      Left            =   1680
      TabIndex        =   15
      Text            =   "Employee address"
      Top             =   2580
      Width           =   4455
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
      Left            =   7380
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Create emp"
      Top             =   60
      Width           =   1575
   End
   Begin TDBDate6Ctl.TDBDate txtCreateDT 
      Height          =   315
      Left            =   5580
      TabIndex        =   7
      Top             =   60
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      Calendar        =   "frmEmployee_Detail.frx":000C
      Caption         =   "frmEmployee_Detail.frx":0107
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEmployee_Detail.frx":016A
      Keys            =   "frmEmployee_Detail.frx":0188
      Spin            =   "frmEmployee_Detail.frx":01E6
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy-mm-dd"
      EditMode        =   3
      Enabled         =   0
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
      Caption         =   "放弃"
      Height          =   435
      Left            =   5820
      TabIndex        =   6
      Top             =   4320
      Width           =   1275
   End
   Begin VB.TextBox txtRemark 
      Appearance      =   0  'Flat
      Height          =   555
      Left            =   1680
      TabIndex        =   5
      Text            =   "Remark"
      Top             =   3180
      Width           =   4035
   End
   Begin VB.TextBox txtEmpID 
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Employee iD"
      Top             =   180
      Width           =   2535
   End
   Begin VB.ComboBox cboEmployeeState 
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
      Height          =   330
      ItemData        =   "frmEmployee_Detail.frx":020E
      Left            =   1680
      List            =   "frmEmployee_Detail.frx":0210
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveDetail 
      Caption         =   "保存"
      Height          =   435
      Left            =   2220
      TabIndex        =   2
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox txtEmployeePhone 
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
      Left            =   1680
      TabIndex        =   1
      Text            =   "Employee phone"
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox txtChsName 
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
      Left            =   1680
      TabIndex        =   0
      Text            =   "Employee name"
      Top             =   1380
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "密码"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "英文名"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   540
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "员工状态"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   3900
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "备注信息"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   3300
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "员工住址"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "员工电话"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "中文名"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "员工编号"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmEmployee_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dicLayout As Scripting.Dictionary
Dim mbolLayout As Boolean
Public mID As Long
Private bolisNew As Boolean
Dim mLastIncrementalControl As VB.Control
Dim mdicCity As New Scripting.Dictionary '3级连动菜单中，承上启下的，省份是独立的，不需要纪录，当前城市列表必须要保存，否则当前行政区就无法取得。
Dim bolArrowKeyDownState As Boolean '在用方向键选择增量下拉框里内容的时候，会出发onclick事件，坑爹啊！只好再加个窗体变量控制一下了。
Dim bolOnLoading As Boolean '由于窗体加载数据的时候，也会出发ONCHANGE事件，所以。。。加载的时候，不能把下拉框SHOW出来。

Public Function LoadDetail(ByVal ID As String) As String
    bolOnLoading = True
    bolisNew = False
    mID = ID
    Dim strUrl, strPostData As String
    strPostData = "{""Type"":""LoadEmployeeDetail"",""Fields"":[""EmpID""],""Values"":[""" & ID & """]}"
    strUrl = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strUrl, strPostData)
    Call FillFormTextBox(Me, JSON.Parse(strResult))
    
    '登陆密码部分呢，要做特殊处理
    Dim objTransPRD As clsTransformPWD
    Set objTransPRD = New clsTransformPWD
    Me.txtPWD.Text = objTransPRD.deTransFormPWD(Me.txtPWD.Text)
    Set objTransPRD = Nothing
    
    bolOnLoading = False
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
    
        If isCtlLinkedDB(ctl) Then
            
            SBField.Append """" & Right(ctl.name, Len(ctl.name) - 3) & ""","

            If ctl.name = "txtPWD" Then
                Dim objTransPRD As clsTransformPWD
                Set objTransPRD = New clsTransformPWD
                SBValue.Append """" & SafeJsonField(objTransPRD.TransFormPWD(MakeQueryValue(Right(ctl.name, Len(ctl.name) - 3), ctl.Text, False))) & ""","
            
                Set objTransPRD = Nothing
            Else
                SBValue.Append """" & MakeQueryValue(Right(ctl.name, Len(ctl.name) - 3), ctl.Text, False) & ""","
            End If
        End If
    
    Next
    
    Dim strPostData As String
    Dim strUrl As String
    Dim strFields As String
    Dim strValues As String
    strFields = SBField.toString
    strValues = SBValue.toString

    strPostData = "{""Type"":""EmployeeDetail"",""Fields"":[" & Left(strFields, Len(strFields) - 1) & "],""Values"":[" & Left(strValues, Len(strValues) - 1) & "]}"
    strUrl = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strUrl, strPostData)
    Debug.Print strResult

    If Left(strResult, 1) <> "{" Then '流氓判断法，先把信息提示出来再说了。
        MsgBox "保存失败，请检查字段信息"
    Else
        MsgBox "保存成功"
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    bolisNew = True
    Call InitLayout(Me)
    Call InitTextBox(Me)
    Me.cboEmployeeState.AddItem "在职"
    Me.cboEmployeeState.AddItem "离职"
    
    Call FillCboWithSampleDic(Me.txtEmployeeProvince, gdicLocation.Item("0"))
    Me.txtEmpID.Text = ""

End Sub
'
'Private Sub lstIncremental_Click()
'    If Me.lstIncremental.Text <> "" And (Not bolOnLoading) And (Not bolArrowKeyDownState) Then
'        mLastIncrementalControl.Text = Me.lstIncremental.Text
'        Me.lstIncremental.Visible = False
'    End If
'End Sub
'
'Private Sub txtEmployeeProvince_Change()
'
'    If Me.txtEmployeeProvince.Text <> "" And Not bolOnLoading Then
'        Call txtEmployeeProvince_GotFocus
'        Me.txtEmployeeCity.Text = ""
'        Me.txtEmployeeDistrict.Text = ""
'        'mdicCity.RemoveAll
'    Else
'        Me.lstIncremental.Visible = False
'    End If
'
'End Sub
'
'Private Sub txtEmployeeProvince_GotFocus()
'Me.lstIncremental.Visible = False
'    Set mLastIncrementalControl = Me.txtEmployeeProvince
'    Me.lstIncremental.Visible = QuichSearchbyLocaldic(Me.txtEmployeeProvince, lstIncremental, gdicLocation.Item("0"))
'End Sub
'
'Private Sub txtEmployeeProvince_KeyDown(KeyCode As Integer, Shift As Integer)
'    bolArrowKeyDownState = True
'    If Me.lstIncremental.Visible Then
'
'        Select Case True
'
'            Case KeyCode = vbKeyDown
'
'                If Me.lstIncremental.ListIndex < Me.lstIncremental.ListCount - 1 Then
'                    Me.lstIncremental.ListIndex = Me.lstIncremental.ListIndex + 1
'                End If
'
'            Case KeyCode = vbKeyUp
'
'                If Me.lstIncremental.ListIndex > 0 Then
'                    Me.lstIncremental.ListIndex = Me.lstIncremental.ListIndex - 1
'                End If
'
'            Case KeyCode = vbKeyReturn
'
'                If Me.lstIncremental.ListIndex >= 0 Then
'                    Me.txtEmployeeProvince.Text = Me.lstIncremental.Text
'                    Me.txtEmployeeCity.Text = ""
'                    Me.txtEmployeeDistrict.Text = ""
'                    mdicCity.RemoveAll
'                    Me.lstIncremental.Visible = False
'                End If
'
'        End Select
'
'    End If
'    bolArrowKeyDownState = False
'End Sub
'
''===========================================================================
'
'Private Sub txtEmployeeCity_Change()
'    If Me.txtEmployeeProvince.Text <> "" And (Not bolOnLoading) And (Not bolArrowKeyDownState) Then
'        Call txtEmployeeCity_GotFocus
'    Else
'        Me.lstIncremental.Visible = False
'    End If
'End Sub
'
'Private Sub txtEmployeeCity_GotFocus()
'    Me.lstIncremental.Visible = False
'    Set mLastIncrementalControl = Me.txtEmployeeCity
'    Dim dic As Scripting.Dictionary
'    Dim strFatherKey As String
'    strFatherKey = FindDicKeyByValue(Me.txtEmployeeProvince.Text, gdicLocation.Item("0"))
'    Set dic = FindSubAreaByFather(strFatherKey)
'    Set mdicCity = dic
'    Me.lstIncremental.Visible = QuichSearchbyLocaldic(Me.txtEmployeeCity, lstIncremental, dic)
'
'End Sub
'
'Private Sub txtEmployeeCity_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    bolArrowKeyDownState = True
'
'    If Me.lstIncremental.Visible Then
'
'        Select Case True
'
'            Case KeyCode = vbKeyDown
'
'                If Me.lstIncremental.ListIndex < Me.lstIncremental.ListCount - 1 Then
'                    Me.lstIncremental.ListIndex = Me.lstIncremental.ListIndex + 1
'                End If
'
'            Case KeyCode = vbKeyUp
'
'                If Me.lstIncremental.ListIndex > 0 Then
'                    Me.lstIncremental.ListIndex = Me.lstIncremental.ListIndex - 1
'                End If
'
'            Case KeyCode = vbKeyReturn
'
'                If Me.lstIncremental.ListIndex >= 0 Then
'                    Me.txtEmployeeCity.Text = Me.lstIncremental.Text
'                    Me.txtEmployeeDistrict.Text = ""
'                    Me.lstIncremental.Visible = False
'                End If
'
'        End Select
'
'    End If
'
'    bolArrowKeyDownState = False
'End Sub
'
'
'
''===========================================================================
'Private Sub txtEmployeeDistrict_Change()
'
'    If Me.txtEmployeeDistrict.Text <> "" And (Not bolOnLoading) And (Not bolArrowKeyDownState) Then
'        Call txtEmployeeDistrict_GotFocus
'    Else
'        Me.lstIncremental.Visible = False
'    End If
'
'End Sub
'
'Private Sub txtEmployeeDistrict_GotFocus()
'    Me.lstIncremental.Visible = False
'    Set mLastIncrementalControl = Me.txtEmployeeDistrict
'    Dim dic As Scripting.Dictionary
'    Dim strFatherKey As String
'    strFatherKey = FindDicKeyByValue(Me.txtEmployeeCity.Text, mdicCity)
'    Set dic = FindSubAreaByFather(strFatherKey)
'    Me.lstIncremental.Visible = QuichSearchbyLocaldic(Me.txtEmployeeDistrict, lstIncremental, dic)
'End Sub
'
'Private Sub txtEmployeeDistrict_KeyDown(KeyCode As Integer, Shift As Integer)
'bolArrowKeyDownState = True
'    If Me.lstIncremental.Visible Then
'
'        Select Case True
'
'            Case KeyCode = vbKeyDown
'
'                If Me.lstIncremental.ListIndex < Me.lstIncremental.ListCount - 1 Then
'                    Me.lstIncremental.ListIndex = Me.lstIncremental.ListIndex + 1
'                End If
'
'            Case KeyCode = vbKeyUp
'
'                If Me.lstIncremental.ListIndex > 0 Then
'                    Me.lstIncremental.ListIndex = Me.lstIncremental.ListIndex - 1
'                End If
'
'            Case KeyCode = vbKeyReturn
'
'                If Me.lstIncremental.ListIndex >= 0 Then
'                    Me.txtEmployeeDistrict.Text = Me.lstIncremental.Text
'                    Me.lstIncremental.Visible = False
'                End If
'
'        End Select
'
'    End If
'bolArrowKeyDownState = False
'End Sub

