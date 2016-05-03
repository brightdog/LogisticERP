VERSION 5.00
Begin VB.Form frmOrdertoCompany 
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9405
   ControlBox      =   0   'False
   Icon            =   "frmOrdertoCompany.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   9405
   StartUpPosition =   3  '窗口缺省
   Tag             =   "取件入库"
   Begin VB.ComboBox cboWareHouseName 
      Enabled         =   0   'False
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
      Left            =   3660
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   4335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   435
      Left            =   3600
      TabIndex        =   5
      Top             =   6900
      Width           =   1995
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8100
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "加入"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8100
      TabIndex        =   4
      Top             =   1380
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印入库单"
      Height          =   435
      Left            =   7080
      TabIndex        =   8
      Top             =   6900
      Width           =   2235
   End
   Begin VB.TextBox txtExpressNo 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3660
      TabIndex        =   3
      Text            =   "Express no"
      Top             =   1380
      Width           =   4335
   End
   Begin VB.ListBox lstExpressNo 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      Left            =   4800
      TabIndex        =   6
      Top             =   2040
      Width           =   4515
   End
   Begin VB.TextBox txtPickupReceiptNo 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3660
      TabIndex        =   0
      Text            =   "Pickup receipt no"
      Top             =   180
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "当前收件仓库"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   180
      TabIndex        =   11
      Top             =   780
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "当前取件单编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   180
      TabIndex        =   10
      Top             =   180
      Width           =   3135
   End
   Begin VB.Label lblPickupReceiptInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   180
      TabIndex        =   9
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "快递单号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   180
      TabIndex        =   7
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Menu mnuOper 
      Caption         =   "操作"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenDetail 
         Caption         =   "打开详情"
         Index           =   1
      End
      Begin VB.Menu mnuDel 
         Caption         =   "删除"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmOrdertoCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mControlsPOI As Scripting.Dictionary

Private Sub cboWareHouseName_Click()

    If Me.cboWareHouseName.ListIndex > -1 Then
        Me.cmdAdd.Enabled = True
    End If

End Sub

Private Sub cmdAdd_Click()
    
    Dim strPickupReceiptNO As String
    Dim strExpressNO As String
    
    Dim strJSON As String
    
    strJSON = "{"
    strJSON = strJSON & """PickupReceiptNO"":""" & Me.txtPickupReceiptNo.Text & """"
    strJSON = strJSON & ",""ExpressNO"":""" & Me.txtExpressNo.Text & """"
    strJSON = strJSON & ",""WareHouseName"":""" & Me.cboWareHouseName.Text & """"
    strJSON = strJSON & "}"
    Dim strResult As String
    strResult = modPostData_Core.PostData("savesingleexpressnotopickupreceipt.asp", strJSON)
    Debug.Print strResult
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)

    If dicResult.Item("ERR") = "EIDDUPLICATE" Then
    
        MsgBox "快递单号重复，请重新输入！", vbInformation + vbOKOnly
        Exit Sub
    End If

    If dicResult.Item("State") = "NEW" Then
        Me.lstExpressNo.AddItem Me.txtExpressNo.Text
        Me.lstExpressNo.Selected(Me.lstExpressNo.ListCount - 1) = True
        
    Else
    
        Dim i As Integer
        
        For i = 0 To Me.lstExpressNo.ListCount - 1
        
            If Me.lstExpressNo.List(i) = Me.txtExpressNo.Text Then
                Me.lstExpressNo.Selected(i) = True
                Exit For
            End If
        
        Next
    
    End If
    
    Me.txtExpressNo.Text = ""
    
End Sub

Private Sub cmdClose_Click()
    
    Call frmMain.CloseTab(Me.Tag)
    Unload Me
End Sub

Private Sub cmdSearch_Click()

    If Trim(Me.txtPickupReceiptNo.Text) <> "" Then
        Debug.Print "txtPickupReceiptNo_LostFocus:" & Me.txtPickupReceiptNo.Text
        Dim dicPickupReceiptInfo As Scripting.Dictionary
        Set dicPickupReceiptInfo = CheckPickupReceiptInfoByID(Me.txtPickupReceiptNo.Text)

        If dicPickupReceiptInfo.Item("ERR") = Empty Then
            
            Dim strInfo As String
            strInfo = ""
            Dim i As Integer
            
            For i = 1 To dicPickupReceiptInfo.Item("Header").Count
                strInfo = strInfo & dicPickupReceiptInfo.Item("Header")(i) & ":" & vbTab & dicPickupReceiptInfo.Item("Rst")(1)(i) & vbCrLf

                If dicPickupReceiptInfo.Item("Header")(i) = "WarehouseID" Then
                    
                    Dim dicWarehouse As Scripting.Dictionary
                    Set dicWarehouse = GetWarehouseListbyID(dicPickupReceiptInfo.Item("Rst")(1)(i))
                    
                    Call FillComboBoxWithDic(Me.cboWareHouseName, dicWarehouse, "WarehouseName")
                
                End If

            Next
            
            Me.lblPickupReceiptInfo.Caption = strInfo
            
            Dim dicList As Scripting.Dictionary
        
            Set dicList = loadCurrentPickupReceiptsExpressNoList(Me.txtPickupReceiptNo.Text)
        
            Call FillListBoxWithDic(Me.lstExpressNo, dicList, "ExpressNO")
        
            Me.cmdAdd.Enabled = True
        
        Else
            'Me.cboWareHouseName.Enabled = False
            Me.cmdAdd.Enabled = False
            Me.lstExpressNo.Clear
    
        End If
    End If

End Sub

Private Sub Form_Load()
    Set mControlsPOI = GetAllControlsPOI(Me)
    Me.txtPickupReceiptNo.Text = ""
    Me.txtExpressNo.Text = ""
    'Dim dicWarehouse As Scripting.Dictionary
    'Set dicWarehouse = GetWarehouseListbyEmpName(gUSERNAME)
    'Call FillComboBoxWithDic(Me.cboWareHouseName, dicWarehouse, "WarehouseName")
    '反正取件单里有的，所以这里就不用额外选择了，直接带过来就可以了。
End Sub

Private Sub Form_Resize()
    Call ResizeFormControls(Me, mControlsPOI, True)
End Sub

'Private Sub lstExpressNo_DblClick()
'
'    If Me.lstExpressNo.ListIndex >= 0 Then
'        If RemoveExpressNOFromPickupReceipt(Me.txtPickupReceiptNo.Text, Me.lstExpressNo.List(Me.lstExpressNo.ListIndex)) Then
'
'            Me.lstExpressNo.RemoveItem Me.lstExpressNo.ListIndex
'        Else
'            MsgBox "操作失败！"
'        End If
'    End If
'
'End Sub

Private Sub lstExpressNo_DblClick()

    If Me.lstExpressNo.ListIndex >= 0 Then
        Load frmOrder_Detail
        Call frmOrder_Detail.LoadDetailByExpressNO(Me.lstExpressNo.List(Me.lstExpressNo.ListIndex))
        frmOrder_Detail.Show vbModal
        Unload frmOrder_Detail
    End If

End Sub

Private Sub RemoveFromList(ListIndex)

    If Me.lstExpressNo.ListIndex >= 0 Then
        If RemoveExpressNOFromPickupReceipt(Me.txtPickupReceiptNo.Text, Me.lstExpressNo.List(Me.lstExpressNo.ListIndex)) Then
        
            Me.lstExpressNo.RemoveItem Me.lstExpressNo.ListIndex
        Else
            MsgBox "操作失败！"
        End If
    End If

End Sub

Private Sub lstExpressNo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    If Button = vbRightButton Then
        Call MouseClick(0, 0)

        DoEvents

        If Me.lstExpressNo.List(Me.lstExpressNo.ListIndex) <> "" Then
            Debug.Print (x + Me.lstExpressNo.Left + Me.Left) / Screen.TwipsPerPixelX & ":" & (y + Me.lstExpressNo.Height + Me.Height) / Screen.TwipsPerPixelY
            Me.PopupMenu mnuOper
        End If
    End If

End Sub

Private Sub mnuDel_Click(Index As Integer)
    If MsgBox("确定删除当前纪录？", vbOKCancel) = vbOK Then
    
        Debug.Print "执行删除操作"
        Call RemoveFromList(Me.lstExpressNo.ListIndex)
    End If
End Sub

Private Sub mnuOpenDetail_Click(Index As Integer)
    Call lstExpressNo_DblClick
End Sub

Private Function RemoveExpressNOFromPickupReceipt(ByVal PickupReceiptNO As String, ByVal ExpressNO As String) As Boolean
    
    Dim strPickupReceiptNO As String
    Dim strExpressNO As String
    
    Dim strJSON As String
    
    strJSON = "{""PickupReceiptNO"":""" & PickupReceiptNO & """,""ExpressNO"":""" & ExpressNO & """}"
    Dim strResult As String
    strResult = modPostData_Core.PostData("removesingleexpressnofrompickupreceipt.asp", strJSON)
    Debug.Print strResult
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)

    If dicResult.Item("State") = "SUCCESS" Then
        
        RemoveExpressNOFromPickupReceipt = True
    Else
        RemoveExpressNOFromPickupReceipt = False
    
    End If
    
End Function

Private Sub txtExpressNo_GotFocus()
    Me.txtExpressNo.SelStart = 0
    Me.txtExpressNo.SelLength = Len(Me.txtExpressNo.Text)
End Sub

Private Sub txtExpressNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call cmdAdd_Click
    End If

End Sub

Private Sub txtPickupReceiptNo_GotFocus()
    Me.txtPickupReceiptNo.SelStart = 0
    Me.txtPickupReceiptNo.SelLength = Len(Me.txtPickupReceiptNo.Text)
    Debug.Print "txtPickupReceiptNo_GotFocus"
End Sub

Private Sub txtPickupReceiptNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call cmdSearch_Click
    End If

End Sub

Private Sub txtPickupReceiptNo_LostFocus()

    '先检测当前取件单号码是否有效
    '如果有效，则允许在当前取件单中，加入新的订单编号
    '（是否要加一个移除订单的功能呢？万一有人输入错误？？）
    '如果有效，则继续加载当前取件单编号下的所有订单到列表框里去。
    '其余一些界面按钮控制等。
    Debug.Print "txtPickupReceiptNo_LostFocus"
    Call cmdSearch_Click
End Sub

