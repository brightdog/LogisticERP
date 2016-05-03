VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPackageDelivery_Detail 
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10335
   ControlBox      =   0   'False
   Icon            =   "frmPackageDelivery_Detail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   10335
   StartUpPosition =   3  '窗口缺省
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
      Left            =   1350
      TabIndex        =   12
      Tag             =   "frmTransfer"
      Top             =   1200
      Width           =   3300
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Height          =   435
      Left            =   8700
      TabIndex        =   10
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "放弃"
      Height          =   435
      Left            =   5400
      TabIndex        =   9
      Top             =   6120
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   435
      Left            =   2520
      TabIndex        =   8
      Top             =   6180
      Width           =   1995
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
      Height          =   5340
      Left            =   4800
      TabIndex        =   6
      Top             =   720
      Width           =   5475
   End
   Begin VB.TextBox txtExpressNo 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6180
      TabIndex        =   5
      Text            =   "Express no"
      Top             =   120
      Width           =   3135
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
      Left            =   1380
      TabIndex        =   3
      Tag             =   "frmTransfer"
      Top             =   780
      Width           =   1860
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "搜索"
      Height          =   315
      Left            =   3660
      TabIndex        =   2
      Top             =   780
      Width           =   1035
   End
   Begin VB.ListBox lstIncremental 
      Height          =   1680
      Left            =   480
      TabIndex        =   0
      Top             =   6060
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSFlexGridLib.MSFlexGrid grdTransferList 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   1620
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7858
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
      Caption         =   "承运人名称："
      Height          =   240
      Left            =   60
      TabIndex        =   13
      Top             =   1290
      Width           =   1275
   End
   Begin VB.Label lblOutWareHouseReceiptNo 
      Caption         =   "Out ware house receipt no"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "快递单号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   7
      Top             =   180
      Width           =   1395
   End
   Begin VB.Label lblOrderID 
      Caption         =   "派送员编号："
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   870
      Width           =   1275
   End
End
Attribute VB_Name = "frmPackageDelivery_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mControlsPOI As Scripting.Dictionary

Private Sub cmdSearch_Click()
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
