VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTransfer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13275
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
   Icon            =   "frmTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ListBox lstIncremental 
      Height          =   1740
      Left            =   4860
      TabIndex        =   25
      Top             =   660
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   1560
      TabIndex        =   1
      Top             =   540
      Width           =   3300
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "新增"
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   6180
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid grdList 
      Height          =   4335
      Left            =   180
      TabIndex        =   4
      Top             =   1200
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   7646
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
      TabIndex        =   8
      Top             =   6300
      Width           =   7575
      Begin VB.ComboBox cboSkip 
         Height          =   330
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   22
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
         TabIndex        =   21
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
         Index           =   2
         Left            =   2940
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
         Index           =   3
         Left            =   3300
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
         Index           =   4
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   0
         Width           =   555
      End
      Begin VB.Label lblPageInfo 
         Caption         =   "pageInfo"
         Height          =   315
         Left            =   0
         TabIndex        =   23
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "搜索"
      Height          =   435
      Left            =   11580
      TabIndex        =   3
      Top             =   540
      Width           =   1155
   End
   Begin VB.ComboBox cboTransferType 
      Height          =   330
      Left            =   7980
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   585
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
      Left            =   1560
      TabIndex        =   0
      Top             =   90
      Width           =   3300
   End
   Begin VB.Label lblSenderName 
      Caption         =   "承运人名称："
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
      Left            =   270
      TabIndex        =   24
      Top             =   630
      Width           =   1275
   End
   Begin VB.Label lblOrderState 
      Caption         =   "承运人类型："
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
      Left            =   6780
      TabIndex        =   7
      Top             =   630
      Width           =   1275
   End
   Begin VB.Label lblOrderID 
      Caption         =   "承运人编号："
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
      Left            =   270
      TabIndex        =   6
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mPageNum As Long
Const mPageSize As Integer = 20
Private bolcanCboSkipWork As Boolean
Private mLastIncrementalControl As VB.Control

Private Sub cmdNew_Click()
    Load frmTransfer_Detail
    frmTransfer_Detail.txtCreateEmp = gUSERNAME
    frmTransfer_Detail.Show vbModal
    
    Unload frmTransfer_Detail
    Call cmdSearch_Click
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
    Call doSearch
    
End Sub

Public Function doSearch(Optional ByVal PageNum As String = 1) As String
    
    Dim dicParam As Scripting.Dictionary
    Set dicParam = New Scripting.Dictionary

'    If PageNum <= 1 Then
        Dim ctl As VB.Control

        For Each ctl In Me.Controls
        
            Select Case TypeName(ctl)
            
                Case "TextBox", "ComboBox"
                    dicParam.Add ctl.name, ctl.Text
            End Select
    
        Next

'    End If

    Dim dicList As Scripting.Dictionary
    
    Set dicList = SearchPagedList(Me.name, dicParam, mPageSize, PageNum)
    
    Call FillGrid(Me.grdList, dicList)
    bolcanCboSkipWork = False
    Call FillPageNavi(Me, dicList)
    bolcanCboSkipWork = True
End Function
'End Function

Private Sub Form_Load()
    Me.Show

    mPageNum = 1
    Me.cboTransferType.AddItem ""
    Me.cboTransferType.AddItem "我司"
    Me.cboTransferType.AddItem "三方"
    
    'Me.txtCreateDT_From.Text = ""
    'Me.txtCreateDT_To.Text = ""
    grdList.rows = 1
    'Call cmdSearch_Click
End Sub

Private Sub Form_Resize()
    Me.grdList.Top = 1000
    Me.grdList.Left = 50
    Me.grdList.Height = Me.Height - 1000 - 350
    Me.grdList.width = Me.width - 100
    '    Me.txtCreateDT_From.Text = Format(Date, "yyyy-mm-dd")
    '    Me.txtCreateDT_To.Text = Format(Date + 1, "yyyy-mm-dd")
    Me.cmdSearch.Left = Me.width - Me.cmdSearch.width - 100
    Me.picPagging.Top = Me.Height - Me.picPagging.Height
    Me.picPagging.Left = Me.width - Me.picPagging.width - 100
    'Me.cmdNewOrder.Left = Me.width - Me.cmdNewOrder.width - 100 - Me.cmdSearch.width
    Me.cmdNew.Top = Me.Height - Me.picPagging.Height - 50
End Sub

Private Sub grdList_DblClick()

    If Me.grdList.Row >= 1 Then
        Load frmTransfer_Detail
        Call frmTransfer_Detail.LoadDetail(Me.grdList.TextMatrix(Me.grdList.Row, 0))
        frmTransfer_Detail.Show vbModal
        Unload frmTransfer_Detail
    End If

End Sub

Private Sub lstIncremental_DBlClick()

    If Me.lstIncremental.Text <> "" Then
    
        mLastIncrementalControl.Text = Me.lstIncremental.Text
        Me.lstIncremental.Visible = False
    End If

End Sub

Private Sub lstIncremental_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
    
        If Me.lstIncremental.Text <> "" Then
    
            mLastIncrementalControl.Text = Me.lstIncremental.Text
            Me.lstIncremental.Visible = False
        End If
    
    End If

End Sub

Private Sub txtTransferName_Change()

    If Me.txtTransferName.Text <> "" Then
        Set mLastIncrementalControl = Me.txtTransferName
        Dim dicIncremental As Scripting.Dictionary
        lstIncremental.Clear
        Set dicIncremental = IncrementalSearch("tblTransfer", "TransferName", "like", Me.txtTransferName.Text)
        Dim i As Integer

        If dicIncremental.Item("Rst").Count > 0 Then

            For i = 1 To dicIncremental.Item("Rst").Count
        
                lstIncremental.AddItem CStr(dicIncremental.Item("Rst").Item(i).Item(1))
    
            Next

            lstIncremental.Move Me.txtTransferName.Left, Me.txtTransferName.Top + Me.txtTransferName.Height, Me.txtTransferName.width, 2000
            Me.lstIncremental.Selected(0) = True
            Me.lstIncremental.Visible = True
        End If

    Else
        Me.lstIncremental.Visible = False
        mLastIncrementalControl = Empty
    End If

End Sub

Private Sub txtSenderName_KeyDown(KeyCode As Integer, Shift As Integer)

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
                    Me.txtTransferName.Text = Me.lstIncremental.Text
                    Me.lstIncremental.Visible = False
                End If

        End Select

    End If

End Sub

