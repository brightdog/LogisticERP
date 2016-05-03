VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form frmOdrtoPickup_Detail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "取件详情"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12300
   Icon            =   "frmOdrtoPickup_Detail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   12300
   StartUpPosition =   2  '屏幕中心
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
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Create emp"
      Top             =   600
      Width           =   2535
   End
   Begin TDBNumber6Ctl.TDBNumber txtInsurePrice 
      Height          =   375
      Left            =   9540
      TabIndex        =   19
      Top             =   4680
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   661
      Calculator      =   "frmOdrtoPickup_Detail.frx":000C
      Caption         =   "frmOdrtoPickup_Detail.frx":002C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOdrtoPickup_Detail.frx":008F
      Keys            =   "frmOdrtoPickup_Detail.frx":00AD
      Spin            =   "frmOdrtoPickup_Detail.frx":00F7
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,##0.##"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,##0.##"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   1935998981
      MinValueVT      =   995885061
   End
   Begin TDBNumber6Ctl.TDBNumber txtWeight 
      Height          =   315
      Left            =   6360
      TabIndex        =   18
      Top             =   4680
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   556
      Calculator      =   "frmOdrtoPickup_Detail.frx":011F
      Caption         =   "frmOdrtoPickup_Detail.frx":013F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOdrtoPickup_Detail.frx":01A2
      Keys            =   "frmOdrtoPickup_Detail.frx":01C0
      Spin            =   "frmOdrtoPickup_Detail.frx":020A
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   1735327749
      MinValueVT      =   1819541509
   End
   Begin TDBNumber6Ctl.TDBNumber txtPkgNum 
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   4080
      Width           =   795
      _Version        =   65536
      _ExtentX        =   1402
      _ExtentY        =   661
      Calculator      =   "frmOdrtoPickup_Detail.frx":0232
      Caption         =   "frmOdrtoPickup_Detail.frx":0252
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOdrtoPickup_Detail.frx":02B5
      Keys            =   "frmOdrtoPickup_Detail.frx":02D3
      Spin            =   "frmOdrtoPickup_Detail.frx":031D
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0"
      EditMode        =   3
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBDate6Ctl.TDBDate txtCreateDT 
      Height          =   315
      Left            =   8820
      TabIndex        =   16
      Top             =   6000
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      Calendar        =   "frmOdrtoPickup_Detail.frx":0345
      Caption         =   "frmOdrtoPickup_Detail.frx":0440
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOdrtoPickup_Detail.frx":04A3
      Keys            =   "frmOdrtoPickup_Detail.frx":04C1
      Spin            =   "frmOdrtoPickup_Detail.frx":051F
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy/mm/dd"
      EditMode        =   3
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   1
      ForeColor       =   -2147483640
      Format          =   "yyyy/mm/dd"
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
      Text            =   "2014/08/12"
      ValidateMode    =   0
      ValueVT         =   2010185735
      Value           =   41863
      CenturyMode     =   0
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   435
      Left            =   3780
      TabIndex        =   15
      Top             =   6180
      Width           =   1275
   End
   Begin VB.TextBox txtRemark 
      Height          =   555
      Left            =   1560
      TabIndex        =   14
      Text            =   "Remark"
      Top             =   4920
      Width           =   4035
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
      Left            =   5040
      TabIndex        =   13
      Text            =   "Pickup receipt iD"
      Top             =   960
      Width           =   3915
   End
   Begin VB.TextBox txtOrderCode 
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
      TabIndex        =   12
      Text            =   "Order code"
      Top             =   960
      Width           =   3915
   End
   Begin VB.TextBox txtOrderID 
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
      TabIndex        =   11
      Text            =   "Order iD"
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox txtPaymentType 
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
      ItemData        =   "frmOdrtoPickup_Detail.frx":0547
      Left            =   9360
      List            =   "frmOdrtoPickup_Detail.frx":0549
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4020
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveDetail 
      Caption         =   "Save"
      Height          =   435
      Left            =   5460
      TabIndex        =   9
      Top             =   6180
      Width           =   2055
   End
   Begin VB.TextBox txtOtherService 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   480
      TabIndex        =   8
      Text            =   "Other service"
      Top             =   3720
      Width           =   3435
   End
   Begin VB.TextBox txtReceiverAddress 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   6360
      TabIndex        =   7
      Text            =   "Receiver address"
      Top             =   2520
      Width           =   4575
   End
   Begin VB.TextBox txtReceiverPhone 
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
      Left            =   6360
      TabIndex        =   6
      Text            =   "Receiver phone"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtReceiverCompany 
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
      Left            =   6360
      TabIndex        =   5
      Text            =   "Receivercompany"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtReceiverName 
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
      Left            =   6360
      TabIndex        =   4
      Text            =   "Receiver name"
      Top             =   1380
      Width           =   4575
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
      Height          =   1035
      Left            =   480
      TabIndex        =   3
      Text            =   "Sender address"
      Top             =   2520
      Width           =   4575
   End
   Begin VB.TextBox txtSenderPhone 
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
      TabIndex        =   2
      Text            =   "Sender phone"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtSenderCompany 
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
      TabIndex        =   1
      Text            =   "Sender company"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtSenderName 
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
      TabIndex        =   0
      Text            =   "Sender name"
      Top             =   1380
      Width           =   4575
   End
End
Attribute VB_Name = "frmOdrtoPickup_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dicLayout As Scripting.Dictionary
Dim mbolLayout As Boolean
Public mID As Long
Private bolisNew As Boolean

Public Function LoadDetail(ByVal ID As String) As String
    bolisNew = False
    mID = ID
    Dim strUrl, strPostData As String
    strPostData = "data={""Type"":""OrderDetail"",""Fields"":[""ID""],""Values"":[""" & ID & """]}"
    strUrl = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strUrl, strPostData)
    Call FillFormTextBox(Me, strResult)
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
    
        If TypeName(ctl) = "TextBox" And ctl.name <> "txtCreateDT" Then
            
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

    strPostData = "data={""Type"":""Detail"",""Fields"":[" & Left(strFields, Len(strFields) - 1) & "],""Values"":[" & Left(strValues, Len(strValues) - 1) & "]}"
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
    Me.txtPaymentType.AddItem "月结"
    Me.txtPaymentType.AddItem "到付"
    
    
    Me.txtOrderID.Text = ""
    
End Sub

