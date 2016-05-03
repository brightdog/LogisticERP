VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPickupReceipt_OverView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "取件详情"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12300
   Icon            =   "frmPickupReceipt_OverView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   12300
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Height          =   435
      Left            =   10050
      TabIndex        =   2
      Top             =   6150
      Width           =   2055
   End
   Begin VB.TextBox txtWarehouseName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Warehouse name"
      Top             =   4920
      Width           =   10995
   End
   Begin VB.TextBox txtTransferName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Transfer name"
      Top             =   4260
      Width           =   10995
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
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Create emp"
      Top             =   60
      Width           =   1755
   End
   Begin TDBDate6Ctl.TDBDate txtCreateDT 
      Height          =   315
      Left            =   10500
      TabIndex        =   6
      Top             =   60
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      Calendar        =   "frmPickupReceipt_OverView.frx":000C
      Caption         =   "frmPickupReceipt_OverView.frx":0107
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPickupReceipt_OverView.frx":016A
      Keys            =   "frmPickupReceipt_OverView.frx":0188
      Spin            =   "frmPickupReceipt_OverView.frx":01E6
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   435
      Left            =   6720
      TabIndex        =   1
      Top             =   6180
      Width           =   1275
   End
   Begin VB.TextBox txtRemark 
      Height          =   495
      Left            =   1020
      TabIndex        =   5
      Text            =   "Remark"
      Top             =   5460
      Width           =   10995
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Pickup receipt iD"
      Top             =   120
      Width           =   2535
   End
   Begin VB.ComboBox cboPickupState 
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
      ItemData        =   "frmPickupReceipt_OverView.frx":020E
      Left            =   6720
      List            =   "frmPickupReceipt_OverView.frx":0210
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "编辑"
      Height          =   435
      Left            =   3720
      TabIndex        =   0
      Top             =   6180
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid grdList 
      Height          =   3375
      Left            =   60
      TabIndex        =   8
      Top             =   720
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   5953
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
   Begin VB.Label Label3 
      Caption         =   "备注"
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   5640
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "收件仓库"
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   5040
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "取件人"
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   4380
      Width           =   795
   End
End
Attribute VB_Name = "frmPickupReceipt_OverView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dicLayout As Scripting.Dictionary
Dim mbolLayout As Boolean
Public mID As Long
Private bolisNew As Boolean
Public mdicHiddenFields As New Scripting.Dictionary '应该是通用的每个需要用隐藏内容的窗体都有的变量。
'考虑了半天，感觉还是放在窗体里比较靠谱，放全局的话，就当心可能会哪里控制不好，导致数据混乱。

Public Function LoadDetail(ByVal ID As String) As String
    bolisNew = False

    If ID <> "" Then
        mID = ID
        Dim strUrl, strPostData As String
        strPostData = "data={""Type"":""Detail"",""Fields"":[""ID""],""Values"":[""" & ID & """]}"
        strUrl = LCase(Me.name) & ".asp"
        Dim strResult As String
        strResult = PostData(strUrl, strPostData)
        Call FillFormTextBox(Me, JSON.Parse(strResult))
    
        Call LoadSampleListToGrdByID("tblOrder", " And PickupReceiptID = " & ID, Me.grdList)
    End If

End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    Load frmPickupReceipt_Detail
    Call frmPickupReceipt_Detail.LoadDetail(Me.txtPickupReceiptID.Text)
    frmPickupReceipt_Detail.Show vbModal
    Unload frmPickupReceipt_Detail
    Call LoadDetail(Me.txtPickupReceiptID.Text)
End Sub

Private Sub cmdPrint_Click()

    Dim strTransferName As String
    Dim strWarehouseName As String
    
    Dim dicList As Scripting.Dictionary
    Set dicList = New Scripting.Dictionary

    Dim i, j As Integer
    
    For i = 1 To Me.grdList.rows - 1
    
        Dim dicOrderDetail As Scripting.Dictionary
        Set dicOrderDetail = New Scripting.Dictionary
        
        For j = 0 To Me.grdList.Cols - 1
        
            dicOrderDetail.Add Me.grdList.TextMatrix(0, j), Me.grdList.TextMatrix(i, j)
        
        Next

        dicList.Add i - 1, dicOrderDetail

    Next
    
    rptPickupReceipt.txtTransferName.Text = Me.txtTransferName.Text
    rptPickupReceipt.txtWarehouseName.Text = Me.txtWarehouseName.Text
    Set rptPickupReceipt.dicData = dicList
    rptPickupReceipt.Show vbModal

End Sub

Private Sub Form_Load()
    bolisNew = True
    Call InitLayout(Me)
    Call InitTextBox(Me)
    Me.cboPickupState.AddItem "待取件"
    Me.cboPickupState.AddItem "取件中"
    Me.cboPickupState.AddItem "已取件"
    
    Me.txtPickupReceiptID.Text = ""
    
End Sub

