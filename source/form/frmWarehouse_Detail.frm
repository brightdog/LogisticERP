VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form frmWarehouse_Detail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "仓库详情"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "frmWarehouse_Detail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7920
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtWarehouseAddress 
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
      Left            =   1560
      TabIndex        =   5
      Text            =   "Warehouse address"
      Top             =   2580
      Width           =   4455
   End
   Begin VB.ComboBox txtWarehouseProvince 
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   1635
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
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Create emp"
      Top             =   120
      Width           =   1815
   End
   Begin TDBDate6Ctl.TDBDate txtCreateDT 
      Height          =   315
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      Calendar        =   "frmWarehouse_Detail.frx":000C
      Caption         =   "frmWarehouse_Detail.frx":0107
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmWarehouse_Detail.frx":016A
      Keys            =   "frmWarehouse_Detail.frx":0188
      Spin            =   "frmWarehouse_Detail.frx":01E6
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
      Caption         =   "关闭"
      Height          =   435
      Left            =   6300
      TabIndex        =   9
      Top             =   4260
      Width           =   1275
   End
   Begin VB.TextBox txtRemark 
      Height          =   555
      Left            =   1560
      TabIndex        =   6
      Text            =   "Remark"
      Top             =   3120
      Width           =   4035
   End
   Begin VB.TextBox txtWarehouseCode 
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
      Left            =   1560
      TabIndex        =   0
      Text            =   "Warehouse code"
      Top             =   540
      Width           =   3915
   End
   Begin VB.TextBox txtWarehouseID 
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Warehouse iD"
      Top             =   180
      Width           =   2535
   End
   Begin VB.ComboBox cboWarehouseType 
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
      ItemData        =   "frmWarehouse_Detail.frx":020E
      Left            =   1560
      List            =   "frmWarehouse_Detail.frx":0210
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3780
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveDetail 
      Caption         =   "Save"
      Height          =   435
      Left            =   2400
      TabIndex        =   8
      Top             =   4260
      Width           =   2055
   End
   Begin VB.TextBox txtWarehousePhone 
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
      Left            =   1560
      TabIndex        =   2
      Text            =   "Warehouse phone"
      Top             =   1380
      Width           =   3495
   End
   Begin VB.TextBox txtWarehouseContactor 
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
      Left            =   1560
      TabIndex        =   3
      Text            =   "Warehouse contactor"
      Top             =   1740
      Width           =   4575
   End
   Begin VB.TextBox txtWarehouseName 
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
      Left            =   1560
      TabIndex        =   1
      Text            =   "Warehouse name"
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "仓库代码"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "仓库ID"
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "仓库名称"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "仓库电话"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "仓库联系人"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1860
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "仓库地址"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "备注信息"
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   3060
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "仓库类型"
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "frmWarehouse_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dicLayout As Scripting.Dictionary
Dim mbolLayout As Boolean
Public mID As Long
Private bolisNew As Boolean
Dim bolOnLoading As Boolean
Public mdicHiddenFields As Scripting.Dictionary

Public Function LoadDetail(ByVal ID As String) As String
    bolOnLoading = True
    bolisNew = False
    mID = ID
    Dim strUrl, strPostData As String
    strPostData = "data={""Type"":""WarehouseDetail"",""Fields"":[""WarehouseID""],""Values"":[""" & ID & """]}"
    strUrl = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strUrl, strPostData)
    Call FillFormTextBox(Me, JSON.Parse(strResult))
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
            SBValue.Append """" & MakeQueryValue(Right(ctl.name, Len(ctl.name) - 3), ctl.Text, False) & ""","
        End If
    
    Next
    
    Dim strPostData As String
    Dim strUrl As String
    Dim strFields As String
    Dim strValues As String
    strFields = SBField.toString
    strValues = SBValue.toString

    strPostData = "data={""Type"":""WarehouseDetail"",""Fields"":[" & Left(strFields, Len(strFields) - 1) & "],""Values"":[" & Left(strValues, Len(strValues) - 1) & "]}"
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
    Me.cboWarehouseType.AddItem "我司"
    Me.cboWarehouseType.AddItem "三方"
    
    Me.txtWarehouseID.Text = ""
    
    Set mdicHiddenFields = New Scripting.Dictionary
    Call FillCboWithSampleDic(Me.txtWarehouseProvince, gdicLocation.Item("0"))
End Sub
