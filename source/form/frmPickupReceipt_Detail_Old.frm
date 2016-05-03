VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form frmPickupReceipt_Detail_Old 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "取件详情"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13245
   Icon            =   "frmPickupReceipt_Detail_Old.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   13245
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrCloseIncremental 
      Interval        =   500
      Left            =   5160
      Top             =   240
   End
   Begin TabDlg.SSTab SSTBSelectedOrders 
      Height          =   2835
      Left            =   180
      TabIndex        =   35
      Top             =   5940
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   5001
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "已选择订单"
      TabPicture(0)   =   "frmPickupReceipt_Detail_Old.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdSelectedOrders"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid grdSelectedOrders 
         Height          =   2415
         Left            =   60
         TabIndex        =   36
         Top             =   360
         Width           =   12795
         _ExtentX        =   22569
         _ExtentY        =   4260
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
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   435
      Left            =   6900
      TabIndex        =   33
      Top             =   8820
      Width           =   1275
   End
   Begin VB.ListBox lstIncremental 
      Height          =   1680
      Left            =   11640
      TabIndex        =   4
      Top             =   960
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
      TabIndex        =   3
      Text            =   "Create emp"
      Top             =   0
      Width           =   1335
   End
   Begin TDBDate6Ctl.TDBDate txtCreateDT 
      Height          =   315
      Left            =   11940
      TabIndex        =   2
      Top             =   0
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   556
      Calendar        =   "frmPickupReceipt_Detail_Old.frx":0028
      Caption         =   "frmPickupReceipt_Detail_Old.frx":0123
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPickupReceipt_Detail_Old.frx":0186
      Keys            =   "frmPickupReceipt_Detail_Old.frx":01A4
      Spin            =   "frmPickupReceipt_Detail_Old.frx":0202
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
      TabIndex        =   1
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
      ItemData        =   "frmPickupReceipt_Detail_Old.frx":022A
      Left            =   8640
      List            =   "frmPickupReceipt_Detail_Old.frx":022C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1755
   End
   Begin TabDlg.SSTab SSTB 
      Height          =   8955
      Left            =   60
      TabIndex        =   5
      Top             =   420
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   15796
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "订单选择"
      TabPicture(0)   =   "frmPickupReceipt_Detail_Old.frx":022E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblOrderID"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblOrderState"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblOrderDateFrom"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblSenderName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "grdOrderList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraDistrict"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtOrderID"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtOrderState"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCreateDT_From"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCreateDT_To"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdOrderDateFrom"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdOrderDateTo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdSearchOrder"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtSenderName"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdSelectTransfer"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "承运人选择"
      TabPicture(1)   =   "frmPickupReceipt_Detail_Old.frx":024A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "grdTransferList"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtTransferID"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtTransferType"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdSearchTransfer"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtTransferName"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtRemark"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdSelectWarehouse"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "收货仓库"
      TabPicture(2)   =   "frmPickupReceipt_Detail_Old.frx":0266
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "grdWarehouseList"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtWarehouseName"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdSearchWarehouse"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtWarehouseType"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtWarehouseID"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdSaveDetail"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.CommandButton cmdSelectWarehouse 
         Caption         =   "Next"
         Height          =   435
         Left            =   -71280
         TabIndex        =   46
         Top             =   8400
         Width           =   2055
      End
      Begin VB.CommandButton cmdSaveDetail 
         Caption         =   "Save"
         Height          =   435
         Left            =   -70920
         TabIndex        =   45
         Top             =   8400
         Width           =   2055
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
         Left            =   -73530
         TabIndex        =   40
         Tag             =   "frmWarehouseName"
         Top             =   420
         Width           =   3300
      End
      Begin VB.ComboBox txtWarehouseType 
         Height          =   300
         Left            =   -68550
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Tag             =   "frmWarehouseName"
         Top             =   795
         Width           =   1770
      End
      Begin VB.CommandButton cmdSearchWarehouse 
         Caption         =   "搜索"
         Height          =   435
         Left            =   -66060
         TabIndex        =   38
         Top             =   600
         Width           =   1155
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
         Left            =   -73530
         TabIndex        =   37
         Tag             =   "frmWarehouseName"
         Top             =   810
         Width           =   3300
      End
      Begin VB.CommandButton cmdSelectTransfer 
         Caption         =   "Next"
         Height          =   435
         Left            =   3720
         TabIndex        =   34
         Top             =   8400
         Width           =   2055
      End
      Begin VB.TextBox txtRemark 
         Height          =   555
         Left            =   -73080
         TabIndex        =   32
         Text            =   "Remark"
         Top             =   7740
         Width           =   8595
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
         Left            =   -73530
         TabIndex        =   27
         Tag             =   "frmTransfer"
         Top             =   810
         Width           =   3300
      End
      Begin VB.CommandButton cmdSearchTransfer 
         Caption         =   "搜索"
         Height          =   435
         Left            =   -66060
         TabIndex        =   26
         Top             =   600
         Width           =   1155
      End
      Begin VB.ComboBox txtTransferType 
         Height          =   300
         Left            =   -68550
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Tag             =   "frmTransfer"
         Top             =   795
         Width           =   1770
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
         Left            =   -73530
         TabIndex        =   24
         Tag             =   "frmTransfer"
         Top             =   420
         Width           =   3300
      End
      Begin VB.TextBox txtSenderName 
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
         Left            =   1320
         TabIndex        =   18
         Tag             =   "frmOrder"
         Top             =   840
         Width           =   3300
      End
      Begin VB.CommandButton cmdSearchOrder 
         Caption         =   "搜索"
         Height          =   435
         Left            =   9000
         TabIndex        =   17
         Top             =   1320
         Width           =   1155
      End
      Begin VB.CommandButton cmdOrderDateTo 
         Caption         =   "..."
         Height          =   315
         Left            =   9510
         TabIndex        =   16
         Top             =   420
         Width           =   465
      End
      Begin VB.CommandButton cmdOrderDateFrom 
         Caption         =   "..."
         Height          =   315
         Left            =   7410
         TabIndex        =   15
         Top             =   420
         Width           =   465
      End
      Begin VB.TextBox txtCreateDT_To 
         Height          =   315
         Left            =   8235
         TabIndex        =   14
         Tag             =   "frmOrder"
         Text            =   "2014-08-02"
         Top             =   420
         Width           =   1290
      End
      Begin VB.TextBox txtCreateDT_From 
         Height          =   315
         Left            =   6135
         TabIndex        =   13
         Tag             =   "frmOrder"
         Text            =   "2014-08-01"
         Top             =   420
         Width           =   1290
      End
      Begin VB.ComboBox txtOrderState 
         Height          =   300
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "frmOrder"
         Top             =   840
         Width           =   1770
      End
      Begin VB.TextBox txtOrderID 
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
         Left            =   1320
         TabIndex        =   11
         Tag             =   "frmOrder"
         Top             =   480
         Width           =   3300
      End
      Begin VB.Frame fraDistrict 
         Caption         =   "地址"
         Height          =   675
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   8835
         Begin VB.TextBox txtSenderProvince 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Tag             =   "frmOrder"
            Text            =   "Sender province"
            Top             =   240
            Width           =   1635
         End
         Begin VB.TextBox txtSenderCity 
            Height          =   315
            Left            =   1860
            TabIndex        =   9
            Tag             =   "frmOrder"
            Text            =   "Sender city"
            Top             =   240
            Width           =   1635
         End
         Begin VB.TextBox txtSenderDistrict 
            Height          =   315
            Left            =   3600
            TabIndex        =   8
            Tag             =   "frmOrder"
            Text            =   "Sender district"
            Top             =   240
            Width           =   1635
         End
         Begin VB.TextBox txtSenderAddress 
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
            Left            =   5400
            TabIndex        =   7
            Tag             =   "frmOrder"
            Text            =   "Sender address"
            Top             =   240
            Width           =   3315
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdOrderList 
         Height          =   3495
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   12915
         _ExtentX        =   22781
         _ExtentY        =   6165
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
      Begin MSFlexGridLib.MSFlexGrid grdTransferList 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   28
         Top             =   1200
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   7435
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
         Height          =   4215
         Left            =   -74880
         TabIndex        =   41
         Top             =   1200
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   7435
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
      Begin VB.Label Label6 
         Caption         =   "仓库编号："
         Height          =   240
         Left            =   -74820
         TabIndex        =   44
         Top             =   510
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "承运人状态："
         Height          =   240
         Left            =   -69750
         TabIndex        =   43
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "仓库名称："
         Height          =   240
         Left            =   -74820
         TabIndex        =   42
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "承运人名称："
         Height          =   240
         Left            =   -74820
         TabIndex        =   31
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "承运人状态："
         Height          =   240
         Left            =   -69750
         TabIndex        =   30
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "承运人编号："
         Height          =   240
         Left            =   -74820
         TabIndex        =   29
         Top             =   510
         Width           =   1275
      End
      Begin VB.Label lblSenderName 
         Caption         =   "客户名称："
         Height          =   240
         Left            =   210
         TabIndex        =   23
         Top             =   930
         Width           =   1275
      End
      Begin VB.Label lblOrderDateFrom 
         Caption         =   "订单录入日期："
         Height          =   240
         Left            =   4905
         TabIndex        =   22
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label lblOrderState 
         Caption         =   "订单状态："
         Height          =   240
         Left            =   4920
         TabIndex        =   21
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblOrderID 
         Caption         =   "订单编号："
         Height          =   240
         Left            =   210
         TabIndex        =   20
         Top             =   540
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmPickupReceipt_Detail_old"
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
    strPostData = "data={""Type"":""LoadPickupReceipt_Detail"",""Fields"":[""PickupReceiptID""],""Values"":[""" & ID & """]}"
    strUrl = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strUrl, strPostData)
    Call FillFormTextBox(Me, strResult)
    Call LoadSampleListToGrdByID("tblOrder", " And PickupReceiptID = " & ID, Me.grdSelectedOrders)

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

    strPostData = "data={""Type"":""PickupReceipt_Detail"",""Fields"":[" & Left(strFields, Len(strFields) - 1) & "],""Values"":[" & Left(strValues, Len(strValues) - 1) & "]}"
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

Private Sub cmdSearchOrder_Click()
   Call doSearch("frmOrder", grdOrderList, , " And PickupReceiptID < 1 ")
    Me.grdSelectedOrders.Cols = Me.grdOrderList.Cols
    
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

Private Sub cmdSelectTransfer_Click()
    If Me.grdSelectedOrders.rows > 1 Then
    
        Me.SSTB.Tab = 1
    
    Else
    
    
    End If
End Sub

Private Sub cmdSelectWarehouse_Click()
    If Me.grdTransferList.RowSel > 0 Then
    
        Me.SSTB.Tab = 2
    
    Else
    
        MsgBox "请先选择承运人"
    End If
End Sub

Private Sub cmdSearchWarehouse_Click()
Call doSearch("frmWarehouse", grdWarehouseList)
End Sub

Private Sub Form_Load()
    bolisNew = True
    Set mdicSelectedOrder = New Scripting.Dictionary
    Call InitLayout(Me)
    Call InitTextBox(Me)

    Me.txtPickupState.AddItem "待取件"
    Me.txtPickupState.AddItem "取件中"
    Me.txtPickupState.AddItem "已取件"
    
    Me.txtPickupReceiptID.Text = ""
    Me.grdSelectedOrders.rows = 1
    Me.SSTB.Tab = 0
End Sub

Private Sub grdOrderList_DblClick()

    '需要将当前行里的内容，进行模块级暂存。
    '防止操作失误，误添加之后，再回滚
    '回滚之后的内容，似乎没有办法插入到原始位置
    '暂时干脆放在表格的最后面APPEND上去算了，反正也是不需要添加的。
    If Me.grdOrderList.Row > 0 Then
        Dim grdIndex As Integer
        grdIndex = grdOrderList.Row
        'mdicSelectedOrder.Add Me.grdOrderList.TextMatrix(Me.grdOrderList.Row, 0), GetGridRowData(Me.grdOrderList, Me.grdOrderList.Row)
        Call Me.grdSelectedOrders.AddItem(GetGridRowData(Me.grdOrderList, Me.grdOrderList.Row), Me.grdSelectedOrders.rows)

        If grdIndex = (Me.grdOrderList.rows - 1) Then
            Me.grdOrderList.rows = grdIndex
        Else
            Me.grdOrderList.RemoveItem (grdIndex)
        End If
    End If

End Sub

Private Sub grdSelectedOrders_DblClick()

    If Me.grdSelectedOrders.Row > 0 Then
        Dim grdIndex As Integer
        grdIndex = grdSelectedOrders.Row
        'mdicSelectedOrder.Add Me.grdOrderList.TextMatrix(Me.grdOrderList.Row, 0), GetGridRowData(Me.grdOrderList, Me.grdOrderList.Row)
        Call Me.grdOrderList.AddItem(GetGridRowData(Me.grdSelectedOrders, Me.grdSelectedOrders.Row), Me.grdOrderList.rows)

        If grdIndex = (Me.grdSelectedOrders.rows - 1) Then
            Me.grdSelectedOrders.rows = grdIndex
        Else
            Me.grdSelectedOrders.RemoveItem (grdIndex)
        End If
    End If

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

Private Sub txtSenderProvince_Change()

    If Me.txtSenderProvince.Text <> "" And Not bolOnLoading Then
        Call modSleep.MySleep(0.5)
        Call txtSenderProvince_GotFocus
        Me.txtSenderCity.Text = ""
        Me.txtSenderDistrict.Text = ""
        'mdicCity.RemoveAll
    Else
        Me.lstIncremental.Visible = False
    End If

End Sub

Private Sub txtSenderProvince_GotFocus()
Me.lstIncremental.Visible = False
    Set mLastIncrementalControl = Me.txtSenderProvince
    Me.lstIncremental.Visible = QuichSearchbyLocaldic(Me.txtSenderProvince, lstIncremental, gdicLocation.Item("0"))
End Sub

Private Sub txtSenderProvince_KeyDown(KeyCode As Integer, Shift As Integer)
    bolArrowKeyDownState = True
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
                    Me.txtSenderProvince.Text = Me.lstIncremental.Text
                    Me.txtSenderCity.Text = ""
                    Me.txtSenderDistrict.Text = ""
                    mdicCity.RemoveAll
                    Me.lstIncremental.Visible = False
                End If

        End Select

    End If
    bolArrowKeyDownState = False
End Sub

'===========================================================================

Private Sub txtSenderCity_Change()
    If Me.txtSenderProvince.Text <> "" And (Not bolOnLoading) And (Not bolArrowKeyDownState) Then
        Call txtSenderCity_GotFocus
    Else
        Me.lstIncremental.Visible = False
    End If
End Sub

Private Sub txtSenderCity_GotFocus()
    Me.lstIncremental.Visible = False
    Set mLastIncrementalControl = Me.txtSenderCity
    Dim dic As Scripting.Dictionary
    Dim strFatherKey As String
    strFatherKey = FindDicKeyByValue(Me.txtSenderProvince.Text, gdicLocation.Item("0"))
    Set dic = FindSubAreaByFather(strFatherKey)
    Set mdicCity = dic
    Me.lstIncremental.Visible = QuichSearchbyLocaldic(Me.txtSenderCity, lstIncremental, dic)

End Sub

Private Sub txtSenderCity_KeyDown(KeyCode As Integer, Shift As Integer)

    bolArrowKeyDownState = True

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
                    Me.txtSenderCity.Text = Me.lstIncremental.Text
                    Me.txtSenderDistrict.Text = ""
                    Me.lstIncremental.Visible = False
                End If

        End Select

    End If

    bolArrowKeyDownState = False
End Sub



'===========================================================================
Private Sub txtSenderDistrict_Change()

    If Me.txtSenderDistrict.Text <> "" And (Not bolOnLoading) And (Not bolArrowKeyDownState) Then
        Call txtSenderDistrict_GotFocus
    Else
        Me.lstIncremental.Visible = False
    End If

End Sub

Private Sub txtSenderDistrict_GotFocus()
    Me.lstIncremental.Visible = False
    Set mLastIncrementalControl = Me.txtSenderDistrict
    Dim dic As Scripting.Dictionary
    Dim strFatherKey As String
    strFatherKey = FindDicKeyByValue(Me.txtSenderCity.Text, mdicCity)
    Set dic = FindSubAreaByFather(strFatherKey)
    Me.lstIncremental.Visible = QuichSearchbyLocaldic(Me.txtSenderDistrict, lstIncremental, dic)
End Sub

Private Sub txtSenderDistrict_KeyDown(KeyCode As Integer, Shift As Integer)
bolArrowKeyDownState = True
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
                    Me.txtSenderDistrict.Text = Me.lstIncremental.Text
                    Me.lstIncremental.Visible = False
                End If

        End Select

    End If
bolArrowKeyDownState = False
End Sub

