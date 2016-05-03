VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOrder 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15210
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Visible         =   0   'False
   Begin VB.TextBox txtThdOrderOutWorkNO 
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
      Left            =   11280
      TabIndex        =   43
      Text            =   "Thd order out work nO"
      Top             =   420
      Width           =   2280
   End
   Begin VB.TextBox txtThdPkgNO 
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
      Left            =   11280
      TabIndex        =   41
      Text            =   "Thd pkg NO"
      Top             =   60
      Width           =   2280
   End
   Begin VB.CommandButton cmdImportOrderDataFromClipBoard 
      Caption         =   "从剪贴版导入EXCEL内容"
      Height          =   375
      Left            =   2700
      TabIndex        =   40
      Top             =   6180
      Width           =   2835
   End
   Begin VB.CommandButton cmdUpdateThirdPartState 
      Caption         =   "更新本页三方物流信息"
      Enabled         =   0   'False
      Height          =   330
      Left            =   9420
      TabIndex        =   39
      Top             =   840
      Width           =   2715
   End
   Begin VB.CheckBox chkThirdPartExpressNUM_BolFrom 
      BackColor       =   &H0000C0C0&
      Caption         =   "仅显示第三方运单"
      Height          =   315
      Left            =   7260
      TabIndex        =   38
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtCustCode 
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
      Left            =   8700
      TabIndex        =   36
      Top             =   480
      Width           =   1020
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "导出当前结果"
      Height          =   375
      Left            =   1200
      TabIndex        =   35
      Top             =   6180
      Width           =   1515
   End
   Begin VB.TextBox txtExpressNO 
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
      Left            =   1440
      TabIndex        =   33
      Text            =   "Express no"
      Top             =   840
      Width           =   2280
   End
   Begin VB.ListBox lstIncremental 
      Height          =   1740
      Left            =   10200
      TabIndex        =   31
      Top             =   2820
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtPickupReceiptID 
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
      Left            =   1440
      TabIndex        =   3
      Text            =   "Pickup receipt no"
      Top             =   480
      Width           =   2280
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
      Left            =   5280
      TabIndex        =   2
      Top             =   480
      Width           =   2340
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "新增"
      Height          =   375
      Left            =   60
      TabIndex        =   29
      Top             =   6180
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid grdList 
      Height          =   4335
      Left            =   180
      TabIndex        =   28
      Top             =   1260
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   7646
      _Version        =   393216
      RowHeightMin    =   350
      AllowBigSelection=   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
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
   Begin VB.PictureBox picPagging 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   7575
      TabIndex        =   12
      Top             =   6300
      Width           =   7575
      Begin VB.ComboBox cboSkip 
         Height          =   330
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdPagging 
         BackColor       =   &H80000009&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3300
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   4740
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   5100
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPaggingFirst 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   16
         Tag             =   "1"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdPaggingPrev 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdPaggingNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5460
         TabIndex        =   14
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdPaggingLast 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5820
         TabIndex        =   13
         Top             =   0
         Width           =   555
      End
      Begin VB.Label lblPageInfo 
         Caption         =   "pageInfo"
         Height          =   315
         Left            =   0
         TabIndex        =   27
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "搜索"
      Default         =   -1  'True
      Height          =   435
      Left            =   12420
      TabIndex        =   7
      Top             =   720
      Width           =   1155
   End
   Begin VB.CommandButton cmdOrderDateTo 
      Caption         =   "..."
      Height          =   315
      Left            =   8670
      TabIndex        =   11
      Top             =   120
      Width           =   465
   End
   Begin VB.CommandButton cmdOrderDateFrom 
      Caption         =   "..."
      Height          =   315
      Left            =   6570
      TabIndex        =   10
      Top             =   120
      Width           =   465
   End
   Begin VB.TextBox txtCreateDT_To 
      Height          =   315
      Left            =   7395
      TabIndex        =   5
      Text            =   "2014-08-02"
      Top             =   120
      Width           =   1290
   End
   Begin VB.TextBox txtCreateDT_From 
      Height          =   315
      Left            =   5295
      TabIndex        =   4
      Text            =   "2014-08-01"
      Top             =   120
      Width           =   1290
   End
   Begin VB.ComboBox cboInventoryState 
      Height          =   330
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   6
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
      Left            =   1440
      TabIndex        =   1
      Top             =   90
      Width           =   2280
   End
   Begin VB.Label lblThdOrderOutWorkNO 
      Caption         =   "三方出货单号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9960
      TabIndex        =   44
      Top             =   510
      Width           =   1275
   End
   Begin VB.Label tblThdPkgNO 
      Caption         =   "三方包装单号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9960
      TabIndex        =   42
      Top             =   150
      Width           =   1275
   End
   Begin VB.Label lblCustCode 
      Caption         =   "月结编号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7740
      TabIndex        =   37
      Top             =   570
      Width           =   1155
   End
   Begin VB.Label lblExpressNO 
      Caption         =   "运单编号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   34
      Top             =   930
      Width           =   1275
   End
   Begin VB.Label lblPickupReceiptNo 
      Caption         =   "取件单编号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   32
      Top             =   570
      Width           =   1275
   End
   Begin VB.Label lblSenderName 
      Caption         =   "客户名称："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3750
      TabIndex        =   30
      Top             =   570
      Width           =   1275
   End
   Begin VB.Label lblOrderDateFrom 
      Caption         =   "订单录入日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3765
      TabIndex        =   9
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label lblOrderState 
      Caption         =   "订单状态："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3750
      TabIndex        =   8
      Top             =   930
      Width           =   1275
   End
   Begin VB.Label lblOrderID 
      Caption         =   "订单编号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mPageNum As Long
Const mPageSize As Integer = 20
Private bolcanCboSkipWork As Boolean
Private mLastIncrementalControl As VB.Control

Private Sub cmdExport_Click()
    'MsgBox "功能增加中，敬请期待~"
    Dim arrOrderFieldHeader() As String
    Dim arrOrderTableFieldHeader() As String
    Dim arrExpressFieldHeader() As String
        
    arrOrderFieldHeader = Split(gstrOrderFieldHeader, "|", -1, vbBinaryCompare)
        
    arrOrderTableFieldHeader = Split(gstrOrderTableFieldHeader, "|", -1, vbBinaryCompare)
        
    arrExpressFieldHeader = Split(gstrExpressFieldHeader, "|", -1, vbBinaryCompare)
        
    Dim dicParam As Scripting.Dictionary
    Set dicParam = MakeSearchParam(Me)
    
    Dim dicList As Scripting.Dictionary
    
    Set dicList = SearchPagedList("frmorder_export", dicParam, 5000, 1) '暂时写死最多导出5000条数据。应该够用了吧？！
    
    If Not dicList Is Nothing Then
        Dim v As Variant
    
        If dicList.Item("RsCount") > 0 Then
    
            Dim i As Integer
            Dim SB As clsStringBuilder
            Set SB = New clsStringBuilder
            Dim dicField As Scripting.Dictionary
            Set dicField = New Scripting.Dictionary
            
            For i = 1 To dicList.Item("Header").Count - 2 '配合数据库里最后2列字段，CreateDT和OrderID不要输出

                'Dim strMappingField As String
                'strMappingField = GetMappingField(dicList.Item("Header")(i), arrOrderTableFieldHeader, arrOrderFieldHeader)
                If VBA.InStr(1, gstrOrderFieldHeader & "|" & gstrExpressFieldHeader, dicList.Item("Header")(i), vbBinaryCompare) > 0 Then
                    
                    If Not dicField.Exists(dicList.Item("Header")(i)) Then
                        SB.Append dicList.Item("Header")(i) & vbTab
                        Call dicField.Add(dicList.Item("Header")(i), i)
                    End If
                End If
        
            Next

            SB.Append vbCrLf

            For i = 1 To dicList.Item("Rst").Count

                
                For Each v In dicField.keys
                
                    SB.Append Replace(dicList.Item("Rst")(i)(dicField.Item(v)), vbTab, " ") & vbTab
                Next


                SB.Append vbCrLf
            Next
        
            Dim Fso As Scripting.FileSystemObject
            Set Fso = New Scripting.FileSystemObject
        
            Dim TS As Scripting.TextStream
        
            If Not Fso.FolderExists(App.path & "\ExportData") Then
                Call Fso.CreateFolder(App.path & "\ExportData")
            End If

            Dim strFilePath As String
            strFilePath = App.path & "\ExportData\ExportData_" & Format(VBA.Now(), "yyyy-mm-dd_hh_mm_ss") & ".csv"
            Set TS = Fso.CreateTextFile(strFilePath, True, True)

            Call TS.Write(SB.toString)

            TS.Close

            Set TS = Nothing
        
            Set Fso = Nothing
            
            If MsgBox("数据导出成功，共计：" & dicList.Item("RsCount") & " 条纪录。数据存放在：" & vbCrLf & strFilePath & vbCrLf & "是否现在打开文件察看？", vbYesNo, "数据导出结果操作提示") = vbYes Then
            
                OpenFileWithSysProgram strFilePath
            
            End If
            
        Else
            MsgBox "没有找到纪录，请重新设置搜索条件。"
        End If

    Else
        MsgBox "其他错误。请联系开发人员"
    End If
    
End Sub

Private Sub cmdImportOrderDataFromClipBoard_Click()
    Call ImportOrderDataFromClipBoard

End Sub

Private Sub ImportOrderDataFromClipBoard()
    Dim strRawData As String
    strRawData = modReadClipBoardData.GetRAWdataFromClipBrd
    Dim strResult As String
    strResult = ImportOrderDataToDB(strRawData)
    
    If strResult <> "" Then
    
        Dim JsonResult As Scripting.Dictionary
    
        Set JsonResult = JSON.Parse(strResult)

        If Not JsonResult Is Nothing Then
        
            If JsonResult.Item("STATE") = "SUCCESS" Then
            
                MsgBox "导入成功"
                '视情况看是否需要自动刷新列表显示
        
            Else
    
                MsgBox JsonResult.Item("DESC")
    
            End If

        Else
            MsgBox "导入有异常"
        End If
    End If

End Sub

Private Sub cmdNew_Click()
    Load frmOrder_Detail
    frmOrder_Detail.txtCreateEmp = gUSERNAME
    frmOrder_Detail.Show vbModal
    
    Unload frmOrder_Detail
    Call cmdSearch_Click
End Sub

Private Sub cmdOrderDateFrom_Click()
    Call Load(frmCalender)
    Set frmCalender.CallerControl = Me.cmdOrderDateFrom
    Set frmCalender.ValueReturnControl = Me.txtCreateDT_From
    frmCalender.Top = Me.txtCreateDT_From.Top + Me.Top + Me.txtCreateDT_From.Height + frmMain.Top + 500
    frmCalender.Left = Me.txtCreateDT_From.Left + Me.Left + frmMain.Top
    Debug.Print frmCalender.Top & ":" & frmCalender.Left
    frmCalender.tdCld.Value = Me.txtCreateDT_From.Text
    frmCalender.Show
End Sub

Private Sub cmdOrderDateTo_Click()
    Call Load(frmCalender)
    Set frmCalender.CallerControl = Me.cmdOrderDateTo
    Set frmCalender.ValueReturnControl = Me.txtCreateDT_To

    frmCalender.Top = Me.txtCreateDT_To.Top + Me.Top + Me.txtCreateDT_To.Height + gTop + 500
    frmCalender.Left = Me.txtCreateDT_To.Left + Me.Left + gLeft
    frmCalender.tdCld.Value = Me.txtCreateDT_To.Text
    frmCalender.Show
End Sub

Private Sub cmdPagging_Click(Index As Integer)
    Call doSearch(Me.cmdPagging(Index).Tag)
End Sub

Private Sub cmdPaggingFirst_Click()
    Call doSearch(1)
End Sub

Private Sub cmdPaggingLast_Click()
    Call doSearch(Me.cmdPaggingLast.Tag)
End Sub

Private Sub cmdPaggingNext_Click()
    Call doSearch(Me.cmdPaggingNext.Tag)
End Sub

Private Sub cmdPaggingPrev_Click()
    Call doSearch(Me.cmdPaggingPrev.Tag)
End Sub

Private Sub cboSkip_Click()

    If bolcanCboSkipWork Then
        Call doSearch(Me.cboSkip.Text)
    End If

End Sub

Private Sub cmdSearch_Click()

    If IsDate(Me.txtCreateDT_From.Text) And IsDate(Me.txtCreateDT_To.Text) Then
        If DateDiff("d", Me.txtCreateDT_From.Text, Me.txtCreateDT_To.Text) > 30 Then
            MsgBox "日期范围最多30天，请重新选择"
            Exit Sub
        End If
    End If

    Call doSearch
    Me.lstIncremental.Visible = False

    If Me.grdList.rows > 1 Then ' And Me.chkThirdPartExpressNUM_BolFrom.Value = 1
        Me.cmdUpdateThirdPartState.Enabled = True
    Else
        Me.cmdUpdateThirdPartState.Enabled = False
    End If

End Sub

Public Function doSearch(Optional ByVal PageNum As String = 1) As String
    
    Dim dicParam As Scripting.Dictionary
    'Set dicParam = New Scripting.Dictionary
    
    Set dicParam = MakeSearchParam(Me)
'    If PageNum <= 1 Then
'        Dim ctl As VB.Control
'
'        For Each ctl In Me.Controls
'
'            Select Case TypeName(ctl)
'
'                Case "TextBox", "ComboBox"
'                    dicParam.Add ctl.name, ctl.Text
'                Case "CheckBox"
'                    dicParam.Add ctl.name, ctl.Value
'            End Select
'
'        Next

'    End If
    
'    If Me.chkThirdPartExpressNUM_From.Value = 1 Then '由于要拆分父子表，所以这样写就不正确了。
'        dicParam.Add "chkOnlyShowThirdCompany", 1
'    End If
    
    Dim dicList As Scripting.Dictionary
    
    Set dicList = SearchPagedList(Me.name, dicParam, mPageSize, PageNum)
    
    Call FillGrid(Me.grdList, dicList)
    bolcanCboSkipWork = False
    Call FillPageNavi(Me, dicList)
    bolcanCboSkipWork = True
End Function




Private Sub cmdUpdateThirdPartState_Click()

    Dim arrExpressNO() As String
    
    arrExpressNO = GetExpressNOList(Me.grdList, 0)

    If UBound(arrExpressNO) >= 0 Then
        Call UpdateThirdPartExpressInfobyClient(arrExpressNO)

        MsgBox "更新程序后台执行中，请稍候查看结果"
    Else
        MsgBox "本页没有第三方物流订单"
    End If

End Sub



Private Sub Form_Load()
    Me.Show

    mPageNum = 1
    Me.txtCreateDT_From.Text = ""
    Me.txtCreateDT_To.Text = ""
    Me.txtPickupReceiptID.Text = ""
    Me.txtExpressNO.Text = ""
    Me.txtThdOrderOutWorkNO.Text = ""
    Me.txtThdPkgNO.Text = ""
    grdList.rows = 1
    
    Me.cboInventoryState.AddItem ""
    Me.cboInventoryState.AddItem "派件中"
    Me.cboInventoryState.AddItem "已出库"
    Me.cboInventoryState.AddItem "已取件"
    Me.txtCreateDT_To.Text = VBA.Format(CStr(VBA.Date) + " 23:59:59", "yyyy-mm-dd")
    Me.txtCreateDT_From.Text = DateAdd("d", -30, Me.txtCreateDT_To.Text)
    'Call cmdSearch_Click
End Sub

Private Sub Form_Resize()
    Me.grdList.Top = 1200
    Me.grdList.Left = 50
    Me.grdList.Height = Me.Height - 1200 - 350
    Me.grdList.width = Me.width - 100
    '    Me.txtCreateDT_From.Text = Format(Date, "yyyy-mm-dd")
    '    Me.txtCreateDT_To.Text = Format(Date + 1, "yyyy-mm-dd")
    Me.cmdSearch.Left = Me.width - Me.cmdSearch.width - 100
    Me.picPagging.Top = Me.Height - Me.picPagging.Height
    Me.picPagging.Left = Me.width - Me.picPagging.width - 100
    
    Me.cmdNew.Top = Me.Height - Me.picPagging.Height - 50
    Me.cmdExport.Top = Me.cmdNew.Top
    Me.cmdExport.Left = Me.cmdNew.Left + Me.cmdNew.width + 50
    Me.cmdImportOrderDataFromClipBoard.Top = Me.cmdNew.Top
    Me.cmdImportOrderDataFromClipBoard.Left = Me.cmdExport.Left + Me.cmdExport.width + 50
    
End Sub

Private Sub grdList_DblClick()

    If Me.grdList.Row >= 1 Then
        Load frmOrder_Detail
        Call frmOrder_Detail.LoadDetail(Me.grdList.TextMatrix(Me.grdList.Row, 0))
        frmOrder_Detail.Show vbModal
        Unload frmOrder_Detail
    End If

End Sub

Private Sub lstIncremental_DBlClick()

    Call SelectCurrentIncrementalText(Me.lstIncremental)

End Sub

Private Sub lstIncremental_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
    
        Call SelectCurrentIncrementalText(Me.lstIncremental)
    
    End If

End Sub

Private Sub txtPickupReceiptID_Change()
    Call ShowIncrementalSearchList("PickupReceipt", "PickupReceiptID", "=", Me.txtPickupReceiptID, Me.lstIncremental)
End Sub

Private Sub txtPickupReceiptID_KeyDown(KeyCode As Integer, Shift As Integer)
    Call SelectIncrementalResult(KeyCode, Me.txtPickupReceiptID, lstIncremental)
End Sub



Private Sub txtSenderName_Change()

    Call ShowIncrementalSearchList("tblOrder", "SenderName", "like", Me.txtSenderName, Me.lstIncremental)
    
End Sub

Private Sub txtSenderName_KeyDown(KeyCode As Integer, Shift As Integer)

    Call SelectIncrementalResult(KeyCode, Me.txtSenderName, lstIncremental)

End Sub

