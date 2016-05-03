VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form frmPickupReceipt_Detail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "取件详情"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13245
   Icon            =   "frmPickupReceipt_Detail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   13245
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdSaveDetail 
      Caption         =   "保存"
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   7320
      Width           =   1755
   End
   Begin VB.TextBox txtWarehouseName 
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
      Height          =   315
      Left            =   1470
      TabIndex        =   6
      Tag             =   "frmWarehouse"
      Top             =   4110
      Width           =   3300
   End
   Begin VB.CommandButton cmdSearchWarehouse 
      Caption         =   "搜索"
      Height          =   435
      Left            =   8940
      TabIndex        =   8
      Top             =   3900
      Width           =   1155
   End
   Begin VB.ComboBox txtWarehouseType 
      Height          =   300
      Left            =   6450
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "frmWarehouse"
      Top             =   4095
      Width           =   1770
   End
   Begin VB.TextBox txtWarehouseID 
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
      Height          =   315
      Left            =   1470
      TabIndex        =   5
      Tag             =   "frmWarehouse"
      Top             =   3720
      Width           =   3300
   End
   Begin VB.TextBox txtTransferID 
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
      Height          =   315
      Left            =   1410
      TabIndex        =   0
      Tag             =   "frmTransfer"
      Top             =   780
      Width           =   3300
   End
   Begin VB.ComboBox txtTransferType 
      Height          =   300
      Left            =   6390
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "frmTransfer"
      Top             =   1155
      Width           =   1770
   End
   Begin VB.CommandButton cmdSearchTransfer 
      Caption         =   "搜索"
      Height          =   435
      Left            =   8880
      TabIndex        =   3
      Top             =   960
      Width           =   1155
   End
   Begin VB.TextBox txtTransferName 
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
      Height          =   315
      Left            =   1410
      TabIndex        =   1
      Tag             =   "frmTransfer"
      Top             =   1170
      Width           =   3300
   End
   Begin VB.Timer tmrCloseIncremental 
      Interval        =   500
      Left            =   5160
      Top             =   240
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   435
      Left            =   7440
      TabIndex        =   11
      Top             =   7320
      Width           =   1275
   End
   Begin VB.ListBox lstIncremental 
      Height          =   1680
      Left            =   10500
      TabIndex        =   16
      Top             =   7140
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Create emp"
      Top             =   0
      Width           =   1335
   End
   Begin TDBDate6Ctl.TDBDate txtCreateDT 
      Height          =   315
      Left            =   11940
      TabIndex        =   14
      Top             =   0
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   556
      Calendar        =   "frmPickupReceipt_Detail.frx":000C
      Caption         =   "frmPickupReceipt_Detail.frx":0107
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPickupReceipt_Detail.frx":016A
      Keys            =   "frmPickupReceipt_Detail.frx":0188
      Spin            =   "frmPickupReceipt_Detail.frx":01E6
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
      ShowLiterals    =   2
      TabAction       =   0
      Text            =   "2014-08-12"
      ValidateMode    =   0
      ValueVT         =   2010185735
      Value           =   41863
      CenturyMode     =   0
   End
   Begin VB.TextBox txtPickupReceiptID 
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
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Pickup receipt iD"
      Top             =   0
      Width           =   2535
   End
   Begin VB.ComboBox txtPickupState 
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
      ItemData        =   "frmPickupReceipt_Detail.frx":020E
      Left            =   8640
      List            =   "frmPickupReceipt_Detail.frx":0210
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   0
      Width           =   1755
   End
   Begin MSFlexGridLib.MSFlexGrid grdTransferList 
      Height          =   2055
      Left            =   60
      TabIndex        =   4
      Top             =   1560
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   3625
      _Version        =   393216
      RowHeightMin    =   350
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grdWarehouseList 
      Height          =   2535
      Left            =   60
      TabIndex        =   9
      Top             =   4500
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   4471
      _Version        =   393216
      RowHeightMin    =   350
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "仓库名称："
      Height          =   240
      Left            =   180
      TabIndex        =   22
      Top             =   4200
      Width           =   1275
   End
   Begin VB.Label Label5 
      Caption         =   "承运人状态："
      Height          =   240
      Left            =   5250
      TabIndex        =   21
      Top             =   4140
      Width           =   1275
   End
   Begin VB.Label Label6 
      Caption         =   "仓库编号："
      Height          =   240
      Left            =   180
      TabIndex        =   20
      Top             =   3810
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "承运人编号："
      Height          =   240
      Left            =   120
      TabIndex        =   19
      Top             =   870
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "承运人状态："
      Height          =   240
      Left            =   5190
      TabIndex        =   18
      Top             =   1200
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "承运人名称："
      Height          =   240
      Left            =   120
      TabIndex        =   17
      Top             =   1260
      Width           =   1275
   End
End
Attribute VB_Name = "frmPickupReceipt_Detail"
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
Dim mdicSelectedOrder As Scripting.Dictionary
Public mdicHiddenFields As New Scripting.Dictionary  '应该是通用的每个需要用隐藏内容的窗体都有的变量。
'考虑了半天，感觉还是放在窗体里比较靠谱，放全局的话，就当心可能会哪里控制不好，导致数据混乱。

Public Function LoadDetail(ByVal ID As String) As String
    bolOnLoading = True
    bolisNew = False
    mID = ID
    Dim strUrl, strPostData As String
    strPostData = "{""Type"":""LoadPickupReceipt_Detail"",""Fields"":[""PickupReceiptID""],""Values"":[""" & ID & """]}"
    strUrl = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strUrl, strPostData)
    Call FillFormTextBox(Me, JSON.Parse(strResult))
    '    Call LoadSampleListToGrdByID("tblOrder", " And PickupReceiptID = " & ID, Me.grdSelectedOrders)

    If Me.txtTransferID.Text <> "" Then
        Call cmdSearchTransfer_Click
       
    End If

    If Me.txtWarehouseID.Text <> "" Then
        Call cmdSearchWarehouse_Click
    End If

    bolOnLoading = False

    If Me.grdTransferList.rows > 1 Then
        Me.grdTransferList.RowSel = 1
    End If

    If Me.grdWarehouseList.rows > 1 Then
        Me.grdWarehouseList.RowSel = 1
    End If

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
    
        If (TypeName(ctl) = "TextBox" Or TypeName(ctl) = "ComboBox") And ctl.name <> "txtCreateDT" And ctl.Tag = "" Then
            
            SBField.Append """" & Right(ctl.name, Len(ctl.name) - 3) & ""","
            SBValue.Append """" & MakeQueryValue(Right(ctl.name, Len(ctl.name) - 3), ctl.Text, False) & ""","
        End If
        
        If TypeName(ctl) = "MSFlexGrid" Then
            Dim i As Integer
            
            Select Case ctl.name
            
                Case "grdSelectedOrders"
                    Dim strSelectedOrders As String
                    strSelectedOrders = ""

                    For i = 1 To ctl.rows - 1
                    
                        strSelectedOrders = strSelectedOrders & ctl.TextMatrix(i, 0) & "|"
                    
                    Next

                    If strSelectedOrders <> "" Then
                    
                        If Right(strSelectedOrders, 1) = "|" Then
                            strSelectedOrders = Left(strSelectedOrders, Len(strSelectedOrders) - 1)
                        End If

                    Else
                        MsgBox "非法操作:请先添加订单到列表!"
                        Exit Sub
                    End If

                    SBField.Append """" & Right(ctl.name, Len(ctl.name) - 3) & ""","
                    SBValue.Append """" & strSelectedOrders & ""","
                
                Case "grdTransferList"

                    If ctl.RowSel > 0 Then
                        SBField.Append """" & Right(ctl.name, Len(ctl.name) - 3) & ""","
                        SBValue.Append """" & ctl.TextMatrix(ctl.RowSel, 0) & ""","
                    
                    Else
                        MsgBox "非法操作:请先选择承运人!"
                        Exit Sub
                    End If

                Case "grdWarehouseList"

                    If ctl.RowSel > 0 Then
                        SBField.Append """" & Right(ctl.name, Len(ctl.name) - 3) & ""","
                        SBValue.Append """" & ctl.TextMatrix(ctl.RowSel, 0) & ""","
                    
                    Else
                        MsgBox "非法操作:请先选择仓库用来接收货物!"
                        Exit Sub
                    End If

            End Select
        
        End If
        
    Next
    
    Dim strPostData As String
    Dim strUrl As String
    Dim strFields As String
    Dim strValues As String
    strFields = SBField.toString
    strValues = SBValue.toString

    strPostData = "{""Type"":""PickupReceipt_Detail"",""Fields"":[" & Left(strFields, Len(strFields) - 1) & "],""Values"":[" & Left(strValues, Len(strValues) - 1) & "]}"
    strUrl = LCase(Me.name) & ".asp"
    Dim strResult As String
    Debug.Print strUrl
    Debug.Print strPostData
    
    strResult = PostData(strUrl, strPostData)
    Debug.Print strResult

    If Left(strResult, 1) <> "{" Then '流氓判断法，先把信息提示出来再说了。
        MsgBox "保存失败，请检查字段信息"
    Else
        MsgBox "保存成功"
        Unload Me
    End If

End Sub

Private Sub cmdSearchTransfer_Click()
    Call doSearch("frmTransfer", grdTransferList)
End Sub

Public Function doSearch(ByVal TagName As String, ByRef grd As MSFlexGrid, Optional ByVal PageNum As String = 1, Optional ByVal AddtionalQueryString As String = "") As String
    
    Dim dicParam As Scripting.Dictionary
    Set dicParam = New Scripting.Dictionary

    'If PageNum < 1 Then
    Dim ctl As VB.Control

    For Each ctl In Me.Controls
        
        If TypeName(ctl) = "TextBox" Then
            If ctl.Tag = TagName Then
                '因为这里需要搜索2种东西，一个是订单，一个是承运人，所以需要用Tag来区分不同种类的条件筛选框
                dicParam.Add ctl.name, ctl.Text
            End If
        End If
    
    Next

    dicParam.Add "AddtionalQueryString", AddtionalQueryString
    'End If
    
    Dim dicList As Scripting.Dictionary
    
    Set dicList = SearchPagedList(TagName, dicParam, 100, 1)
    
    Call FillGrid(grd, dicList)
    'bolcanCboSkipWork = False
    'Call FillPageNavi(Me, dicList)
    'bolcanCboSkipWork = True
End Function

Private Sub cmdSearchWarehouse_Click()
    Call doSearch("frmWarehouse", grdWarehouseList)
End Sub

Private Sub Form_Load()
    bolisNew = True
    Set mdicSelectedOrder = New Scripting.Dictionary
    Call InitLayout(Me)
    Call InitTextBox(Me)

    Me.txtPickupState.AddItem "待取件"
    Me.txtPickupState.AddItem "已取件"
    Me.txtPickupState.ListIndex = 0
    Me.txtPickupReceiptID.Text = ""

End Sub

Private Sub lstIncremental_Click()

    If Me.lstIncremental.Text <> "" And (Not bolOnLoading) And (Not bolArrowKeyDownState) Then
        mLastIncrementalControl.Text = Me.lstIncremental.Text
        Me.lstIncremental.Visible = False
    End If

End Sub

Private Sub tmrCloseIncremental_Timer()

    If Me.lstIncremental.Visible Then
        If Me.ActiveControl.name <> mLastIncrementalControl.name Then
            Me.lstIncremental.Visible = False
        End If
    End If

End Sub

