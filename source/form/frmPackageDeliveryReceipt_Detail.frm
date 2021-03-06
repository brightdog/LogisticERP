VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form frmPackageDeliveryReceipt_Detail 
   Caption         =   "出库单明细"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   Icon            =   "frmPackageDeliveryReceipt_Detail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   7380
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboWareHouseName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   660
      Width           =   4335
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
      Left            =   1230
      TabIndex        =   2
      Tag             =   "frmTransfer"
      Top             =   2040
      Width           =   3300
   End
   Begin VB.CommandButton cmdShowDetail 
      Caption         =   "详情"
      Height          =   435
      Left            =   5880
      TabIndex        =   7
      Top             =   6180
      Width           =   1455
   End
   Begin VB.TextBox txtPackageDeliveryReceiptID 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   60
      TabIndex        =   13
      Top             =   60
      Width           =   3975
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Create emp"
      Top             =   60
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Height          =   435
      Left            =   60
      TabIndex        =   10
      Top             =   6180
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   435
      Left            =   4440
      TabIndex        =   6
      Top             =   6180
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   435
      Left            =   2220
      TabIndex        =   5
      Top             =   6180
      Width           =   1995
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
      Left            =   1230
      TabIndex        =   1
      Tag             =   "frmTransfer"
      Top             =   1680
      Width           =   1860
   End
   Begin VB.CommandButton cmdSearchTransfer 
      Caption         =   "搜索"
      Height          =   315
      Left            =   6180
      TabIndex        =   3
      Top             =   1680
      Width           =   1035
   End
   Begin VB.ListBox lstIncremental 
      Height          =   1680
      Left            =   480
      TabIndex        =   8
      Top             =   6060
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSFlexGridLib.MSFlexGrid grdTransferList 
      Height          =   3435
      Left            =   0
      TabIndex        =   4
      Top             =   2460
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6059
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
   Begin TDBDate6Ctl.TDBDate txtCreateDT 
      Height          =   315
      Left            =   5940
      TabIndex        =   12
      Top             =   60
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   556
      Calendar        =   "frmPackageDeliveryReceipt_Detail.frx":000C
      Caption         =   "frmPackageDeliveryReceipt_Detail.frx":0107
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPackageDeliveryReceipt_Detail.frx":016A
      Keys            =   "frmPackageDeliveryReceipt_Detail.frx":0188
      Spin            =   "frmPackageDeliveryReceipt_Detail.frx":01E6
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
   Begin VB.Label Label1 
      Caption         =   "出货仓库："
      Height          =   240
      Left            =   60
      TabIndex        =   15
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label lblSenderName 
      Caption         =   "承运人名称："
      Height          =   240
      Left            =   30
      TabIndex        =   14
      Top             =   2130
      Width           =   1275
   End
   Begin VB.Label lblTransferID 
      Caption         =   "派送员编号："
      Height          =   240
      Left            =   30
      TabIndex        =   9
      Top             =   1770
      Width           =   1275
   End
End
Attribute VB_Name = "frmPackageDeliveryReceipt_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mID As Long
Public mdicHiddenFields As Scripting.Dictionary
Dim mControlsPOI As Scripting.Dictionary
Dim bolisNew As Boolean
Dim dicWarehouse As Scripting.Dictionary

Public Function LoadDetail(ByVal ID As String) As String
    bolisNew = False
    mID = ID
    Dim strUrl, strPostData As String
    strPostData = "{""Type"":""LoadPackageDelivery_Detail"",""Fields"":[""PackageDeliveryReceiptID""],""Values"":[""" & ID & """]}"
    strUrl = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strUrl, strPostData)
    Debug.Print strPostData
    Call FillFormTextBox(Me, JSON.Parse(strResult))
    Call cmdSearchTransfer_Click

    Set dicWarehouse = GetWarehouseListbyID(mdicHiddenFields.Item("WarehouseID"))
                    
    Call FillComboBoxWithDic(Me.cboWareHouseName, dicWarehouse, "WarehouseName")
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

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

            If ctl.name = "cboWareHouseName" Then
                SBField.Append """" & "WarehouseID" & ""","
                SBValue.Append """" & MakeQueryValue("WarehouseID", GetWareHouseIDByNameWithSampleDic(dicWarehouse, ctl.Text), False) & ""","
            
            Else
                SBField.Append """" & Right(ctl.name, Len(ctl.name) - 3) & ""","
                SBValue.Append """" & MakeQueryValue(Right(ctl.name, Len(ctl.name) - 3), ctl.Text, False) & ""","
            End If
        End If
        
        If TypeName(ctl) = "MSFlexGrid" Then
            Dim i As Integer
            
            Select Case ctl.name
                
                Case "grdTransferList"

                    If ctl.RowSel > 0 Then
                        SBField.Append """" & Right(ctl.name, Len(ctl.name) - 3) & ""","
                        SBValue.Append """" & ctl.TextMatrix(ctl.RowSel, 0) & ""","
                    ElseIf ctl.rows = 2 Then
                        SBField.Append """" & Right(ctl.name, Len(ctl.name) - 3) & ""","
                        SBValue.Append """" & ctl.TextMatrix(1, 0) & ""","
                        
                    Else
                        MsgBox "非法操作:请先选择派送员!"
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

    strPostData = "{""Type"":""PackageDelivery_Detail"",""Fields"":[" & Left(strFields, Len(strFields) - 1) & "],""Values"":[" & Left(strValues, Len(strValues) - 1) & "]}"
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

Private Sub cmdShowDetail_Click()
    Load frmPackageDelivery
    frmPackageDelivery.txtPackageDeliveryReceiptID.Text = Me.txtPackageDeliveryReceiptID.Text
    Call frmPackageDelivery.Search(Me.txtPackageDeliveryReceiptID.Text)
    frmPackageDelivery.Show vbModal
    Unload frmPackageDelivery
End Sub

Private Sub Form_Load()
    Set mControlsPOI = GetAllControlsPOI(Me)

    '    If IsNumeric(Me.txtPackageDeliveryReceiptID.Text) Then
    '
    '        Call loadCurrentPackageDeliveryReceiptsExpressNoList(Me.txtPackageDeliveryReceiptID.Text)
    '
    '    End If

    bolisNew = True
    
    Set dicWarehouse = GetWarehouseListbyID()
                    
    Call FillComboBoxWithDic(Me.cboWareHouseName, dicWarehouse, "WarehouseName")
    Set mdicHiddenFields = New Scripting.Dictionary
End Sub

Private Sub Form_Resize()
    Call ResizeFormControls(Me, mControlsPOI, True)
End Sub
