VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form frmOrder_Detail 
   AutoRedraw      =   -1  'True
   Caption         =   "订单详情"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14340
   Icon            =   "frmOrder_Detail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   14340
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtReceiverMobile 
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
      Left            =   4500
      TabIndex        =   9
      Text            =   "Receiver mobile"
      Top             =   4740
      Width           =   2055
   End
   Begin VB.TextBox txtSenderMobile 
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
      Left            =   4500
      TabIndex        =   3
      Text            =   "Sender mobile"
      Top             =   2100
      Width           =   2055
   End
   Begin VB.ListBox lstIncremental 
      Height          =   1680
      Left            =   3300
      TabIndex        =   49
      Top             =   8280
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CheckBox chkSpecialHandling_ReturnReceipt 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   "签单返还(托运单)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   11460
      TabIndex        =   29
      Tag             =   "SpecialHandling"
      Top             =   3720
      Width           =   2115
   End
   Begin VB.Timer TmrSetSingleChkbox 
      Interval        =   200
      Left            =   4500
      Top             =   60
   End
   Begin VB.ComboBox cboReceiverProvince 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   11
      Text            =   "Cbo receiver province"
      Top             =   5700
      Width           =   1275
   End
   Begin VB.ComboBox cboSenderProvince 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1020
      TabIndex        =   5
      Text            =   "Cbo sender province"
      Top             =   3000
      Width           =   1275
   End
   Begin VB.CheckBox chkService_Cargo 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   $"frmOrder_Detail.frx":000C
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   11940
      TabIndex        =   22
      Tag             =   "Service"
      Top             =   2220
      Width           =   1395
   End
   Begin VB.CheckBox chkService_Economy 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   $"frmOrder_Detail.frx":0021
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   10320
      TabIndex        =   21
      Tag             =   "Service"
      Top             =   2220
      Width           =   1395
   End
   Begin VB.CheckBox chkService_NextDay 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   $"frmOrder_Detail.frx":0038
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   8760
      TabIndex        =   20
      Tag             =   "Service"
      Top             =   2220
      Width           =   1395
   End
   Begin VB.CheckBox chkService_SameDay 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   $"frmOrder_Detail.frx":004E
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   7200
      TabIndex        =   19
      Tag             =   "Service"
      Top             =   2220
      Width           =   1395
   End
   Begin VB.CheckBox chkPackage_Other 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   $"frmOrder_Detail.frx":0064
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   11940
      TabIndex        =   26
      Tag             =   "Package"
      Top             =   3000
      Width           =   1395
   End
   Begin VB.CheckBox chkPackage_WoodenCase 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   $"frmOrder_Detail.frx":0075
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   10320
      TabIndex        =   25
      Tag             =   "Package"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CheckBox chkPackage_Carton 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   $"frmOrder_Detail.frx":008B
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   8760
      TabIndex        =   24
      Tag             =   "Package"
      Top             =   3000
      Width           =   1395
   End
   Begin VB.CheckBox chkPackage_Envelope 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   $"frmOrder_Detail.frx":009E
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   7260
      TabIndex        =   23
      Tag             =   "Package"
      Top             =   3000
      Width           =   1395
   End
   Begin VB.CheckBox chkSpecialHandling_ReturnFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   "签单往返(客户文件)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   9420
      TabIndex        =   28
      Tag             =   "SpecialHandling"
      Top             =   3720
      Width           =   2115
   End
   Begin VB.CheckBox chkSpecialHandling_SelfTake 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   $"frmOrder_Detail.frx":00B2
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   7260
      TabIndex        =   27
      Tag             =   "SpecialHandling"
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CheckBox chkPaymentType_Third 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   $"frmOrder_Detail.frx":00D3
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   11460
      TabIndex        =   32
      Tag             =   "PaymentType"
      Top             =   4560
      Width           =   1635
   End
   Begin VB.CheckBox chkPaymentType_Cash 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   $"frmOrder_Detail.frx":00F2
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   9420
      TabIndex        =   31
      Tag             =   "PaymentType"
      Top             =   4560
      Width           =   1635
   End
   Begin VB.CheckBox chkPaymentType_Month 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6F4E9&
      Caption         =   $"frmOrder_Detail.frx":0105
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   7260
      TabIndex        =   30
      Tag             =   "PaymentType"
      Top             =   4560
      Width           =   1635
   End
   Begin VB.CommandButton cmdSearchCustCode 
      Caption         =   "搜索"
      Height          =   315
      Left            =   6000
      TabIndex        =   46
      Top             =   1620
      Width           =   795
   End
   Begin VB.TextBox txtCustCode 
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
      Left            =   4080
      TabIndex        =   0
      Text            =   "Cust code"
      Top             =   1620
      Width           =   1815
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
      Left            =   8940
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "Create emp"
      Top             =   0
      Width           =   1995
   End
   Begin TDBNumber6Ctl.TDBNumber txtInsurePrice 
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   7680
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   450
      Calculator      =   "frmOrder_Detail.frx":0119
      Caption         =   "frmOrder_Detail.frx":0139
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":019C
      Keys            =   "frmOrder_Detail.frx":01BA
      Spin            =   "frmOrder_Detail.frx":0204
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
   Begin TDBDate6Ctl.TDBDate txtReceiveDateTime 
      Height          =   315
      Left            =   11280
      TabIndex        =   35
      Top             =   7200
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   556
      Calendar        =   "frmOrder_Detail.frx":022C
      Caption         =   "frmOrder_Detail.frx":0327
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":038A
      Keys            =   "frmOrder_Detail.frx":03A8
      Spin            =   "frmOrder_Detail.frx":0406
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
      EditMode        =   2
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
      ShowContextMenu =   -1
      ShowLiterals    =   2
      TabAction       =   0
      Text            =   ""
      ValidateMode    =   0
      ValueVT         =   2010185729
      Value           =   2.12482986761524E-314
      CenturyMode     =   0
   End
   Begin TDBNumber6Ctl.TDBNumber txtPkgWeight 
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   6900
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      Calculator      =   "frmOrder_Detail.frx":042E
      Caption         =   "frmOrder_Detail.frx":044E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":04B1
      Keys            =   "frmOrder_Detail.frx":04CF
      Spin            =   "frmOrder_Detail.frx":0519
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
      Height          =   315
      Left            =   1380
      TabIndex        =   13
      Top             =   6900
      Width           =   915
      _Version        =   65536
      _ExtentX        =   1614
      _ExtentY        =   556
      Calculator      =   "frmOrder_Detail.frx":0541
      Caption         =   "frmOrder_Detail.frx":0561
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":05C4
      Keys            =   "frmOrder_Detail.frx":05E2
      Spin            =   "frmOrder_Detail.frx":062C
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
      Left            =   10980
      TabIndex        =   44
      Top             =   0
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      Calendar        =   "frmOrder_Detail.frx":0654
      Caption         =   "frmOrder_Detail.frx":074F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":07B2
      Keys            =   "frmOrder_Detail.frx":07D0
      Spin            =   "frmOrder_Detail.frx":082E
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
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "2014-08-12"
      ValidateMode    =   0
      ValueVT         =   2010185735
      Value           =   41863
      CenturyMode     =   0
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "取消"
      Height          =   435
      Left            =   7380
      TabIndex        =   43
      Top             =   8220
      Width           =   1275
   End
   Begin VB.TextBox txtRemark 
      Height          =   315
      Left            =   9180
      TabIndex        =   37
      Text            =   "Remark"
      Top             =   8280
      Width           =   4995
   End
   Begin VB.TextBox txtReceiverSigner 
      Height          =   270
      Left            =   11940
      TabIndex        =   36
      Text            =   "Receiver signer"
      Top             =   6720
      Width           =   1455
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
      Left            =   5640
      TabIndex        =   42
      Text            =   "Pickup receipt iD"
      Top             =   0
      Width           =   2955
   End
   Begin VB.TextBox txtExpressNO 
      Alignment       =   2  'Center
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
      Left            =   7260
      TabIndex        =   41
      Text            =   "Order code"
      Top             =   1380
      Width           =   2955
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "Order iD"
      Top             =   0
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
      ItemData        =   "frmOrder_Detail.frx":0856
      Left            =   12660
      List            =   "frmOrder_Detail.frx":0858
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveDetail 
      Caption         =   "保存"
      Height          =   435
      Left            =   4800
      TabIndex        =   38
      Top             =   8220
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
      Height          =   315
      Left            =   7260
      TabIndex        =   33
      Text            =   "Other service"
      Top             =   5580
      Width           =   6315
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
      Height          =   435
      Left            =   2340
      TabIndex        =   12
      Text            =   "Receiver address"
      Top             =   5640
      Width           =   4395
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
      Left            =   960
      TabIndex        =   8
      Text            =   "Receiver phone"
      Top             =   4740
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
      Left            =   1200
      TabIndex        =   10
      Text            =   "Receivercompany"
      Top             =   5160
      Width           =   5475
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
      Left            =   1440
      TabIndex        =   7
      Text            =   "Receiver name"
      Top             =   4320
      Width           =   4275
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
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Text            =   "Sender address"
      Top             =   2880
      Width           =   4395
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
      Left            =   960
      TabIndex        =   2
      Text            =   "Sender phone"
      Top             =   2100
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
      Left            =   1140
      TabIndex        =   4
      Text            =   "Sender company"
      Top             =   2520
      Width           =   5535
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
      Left            =   1320
      TabIndex        =   1
      Text            =   "Sender name"
      Top             =   1620
      Width           =   1695
   End
   Begin TDBNumber6Ctl.TDBNumber txtPkgLength 
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   7320
      Width           =   555
      _Version        =   65536
      _ExtentX        =   979
      _ExtentY        =   450
      Calculator      =   "frmOrder_Detail.frx":085A
      Caption         =   "frmOrder_Detail.frx":087A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":08DD
      Keys            =   "frmOrder_Detail.frx":08FB
      Spin            =   "frmOrder_Detail.frx":0945
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
   Begin TDBNumber6Ctl.TDBNumber txtPkgWidth 
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   7320
      Width           =   555
      _Version        =   65536
      _ExtentX        =   979
      _ExtentY        =   450
      Calculator      =   "frmOrder_Detail.frx":096D
      Caption         =   "frmOrder_Detail.frx":098D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":09F0
      Keys            =   "frmOrder_Detail.frx":0A0E
      Spin            =   "frmOrder_Detail.frx":0A58
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
   Begin TDBNumber6Ctl.TDBNumber txtPkgHeight 
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   7320
      Width           =   555
      _Version        =   65536
      _ExtentX        =   979
      _ExtentY        =   450
      Calculator      =   "frmOrder_Detail.frx":0A80
      Caption         =   "frmOrder_Detail.frx":0AA0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":0B03
      Keys            =   "frmOrder_Detail.frx":0B21
      Spin            =   "frmOrder_Detail.frx":0B6B
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
   Begin TDBDate6Ctl.TDBDate txtPickupDate 
      Height          =   315
      Left            =   8040
      TabIndex        =   34
      Top             =   7200
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   556
      Calendar        =   "frmOrder_Detail.frx":0B93
      Caption         =   "frmOrder_Detail.frx":0C8E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":0CF1
      Keys            =   "frmOrder_Detail.frx":0D0F
      Spin            =   "frmOrder_Detail.frx":0D6D
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
      EditMode        =   2
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
      ShowContextMenu =   -1
      ShowLiterals    =   2
      TabAction       =   0
      Text            =   ""
      ValidateMode    =   0
      ValueVT         =   2010185729
      Value           =   2.12482986761524E-314
      CenturyMode     =   0
   End
   Begin VB.Label lblSenderCity 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10920
      TabIndex        =   48
      Top             =   1140
      Width           =   1515
   End
   Begin VB.Image imgExpressNO_Barcode 
      Height          =   1110
      Left            =   6960
      Picture         =   "frmOrder_Detail.frx":0D95
      Stretch         =   -1  'True
      Top             =   540
      Width           =   3660
   End
   Begin VB.Label lblReceiverCity 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12540
      TabIndex        =   47
      Top             =   1140
      Width           =   1515
   End
   Begin VB.Image imgBackGround 
      Appearance      =   0  'Flat
      Height          =   7830
      Left            =   0
      Picture         =   "frmOrder_Detail.frx":152AF
      Stretch         =   -1  'True
      Top             =   360
      Width           =   14265
   End
End
Attribute VB_Name = "frmOrder_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dicLayout As Scripting.Dictionary
Dim mbolLayout As Boolean
Public mID As Long
Private bolisNew As Boolean
Public mdicHiddenFields As New Scripting.Dictionary
Dim mControlsPOI As Scripting.Dictionary

Public Function LoadDetail(ByVal ID As String) As String
    bolisNew = False
    mID = ID
    Dim strUrl, strPostData As String
    strPostData = "{""Type"":""OrderDetail"",""Fields"":[""OrderID""],""Values"":[""" & ID & """]}"
    strUrl = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strUrl, strPostData)
    Call FillFormTextBox(Me, JSON.Parse(strResult))
    Me.lstIncremental.Visible = False
End Function

Public Function LoadDetailByExpressNO(ByVal ID As String) As String
    bolisNew = False
    mID = ID
    Dim strUrl, strPostData As String
    strPostData = "{""Type"":""OrderDetailByExpressNO"",""Fields"":[""ExpressNO""],""Values"":[""" & ID & """]}"
    strUrl = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strUrl, strPostData)
    Call FillFormTextBox(Me, JSON.Parse(strResult))
End Function

Private Sub cboReceiverProvince_Change()
    Call cboReceiverProvince_Click
End Sub

Private Sub cboReceiverProvince_Click()
    Me.lblReceiverCity.Caption = Me.cboReceiverProvince.Text
End Sub

Private Sub cboSenderProvince_Change()
    Call cboSenderProvince_Click
End Sub

Private Sub cboSenderProvince_Click()
    Me.lblSenderCity.Caption = Me.cboSenderProvince.Text
End Sub

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

        'Debug.Print "sAVe:" & ctl.name & ":" & ctl.Value
        If isCtlLinkedDB(ctl) Then
            
            SBField.Append """" & Right(ctl.name, Len(ctl.name) - 3) & ""","

            If TypeName(ctl) = "CheckBox" Or TypeName(ctl) = "TDBDate" Then
                SBValue.Append """" & MakeQueryValue(Right(ctl.name, Len(ctl.name) - 3), ctl.Value & "", False) & ""","
            Else
                Debug.Print ctl.name
                SBValue.Append """" & MakeQueryValue(Right(ctl.name, Len(ctl.name) - 3), ctl.Text, False) & ""","
            End If
        End If
    
    Next
    
    Dim strPostData As String
    Dim strUrl As String
    Dim strFields As String
    Dim strValues As String
    strFields = SBField.toString
    strValues = SBValue.toString

    strPostData = "{""Type"":""OrderDetail"",""Fields"":[" & Left(strFields, Len(strFields) - 1) & "],""Values"":[" & Left(strValues, Len(strValues) - 1) & "]}"
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

Private Sub cmdEditLayout_Click()
    Dim ctl As VB.Control
    Set dicLayout = New Scripting.Dictionary

    For Each ctl In Me.Controls
    
        If TypeName(ctl) = "TextBox" Then
            
            ctl.MousePointer = 15
            ctl.Locked = True
            Dim dicSinglePos As Scripting.Dictionary
            Set dicSinglePos = New Scripting.Dictionary
            dicSinglePos.Add "Top", ctl.Top
            dicSinglePos.Add "Left", ctl.Left
            dicSinglePos.Add "X", 0
            dicSinglePos.Add "Y", 0
            dicSinglePos.Add "OffsetX", 0
            dicSinglePos.Add "OffsetY", 0
            
            dicLayout.Add ctl.name, dicSinglePos
            Set dicSinglePos = Nothing
        End If
    
    Next

    mbolLayout = True
    
End Sub

Private Sub cmdSaveLayout_Click()
    Dim ctl As VB.Control
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder

    For Each ctl In Me.Controls
    
        If TypeName(ctl) = "TextBox" Then
            SB.Append """" & ctl.name & """:{"
            SB.Append """Top"":" & """" & ctl.Top & ""","
            SB.Append """Left"":" & """" & ctl.Left & ""","
            SB.Append """Height"":" & """" & ctl.Height & ""","
            SB.Append """Width"":" & """" & ctl.width & """},"
            ctl.MousePointer = 0
            ctl.Locked = True
        End If
    
    Next
    
    Dim strContent As String
    strContent = SB.toString
    strContent = Left(strContent, Len(strContent) - 1)
    
    Dim Fso As Scripting.FileSystemObject
    Set Fso = New Scripting.FileSystemObject
    Dim oText As Scripting.TextStream
    Set oText = Fso.OpenTextFile(APP_CONFIG_PATH & Me.name & ".Layout", ForWriting, True, TristateFalse)
    
    oText.Write "{" & strContent & "}"

    oText.Close
    Set SB = Nothing
    Set Fso = Nothing
    mbolLayout = False
End Sub

Private Sub cmdSearchCustCode_Click()
    '根据客户的月结编号，搜索客户发件人信息，填入下面的文本框中！
    Dim strCustCode As String
    
    strCustCode = Me.txtCustCode.Text

    If strCustCode <> "" Then
    
        Dim dicCust As Scripting.Dictionary
    
        Set dicCust = SearchCustInfoByCode(strCustCode)
        
        '没办法，控件名称无法映射，只能手工赋值，好在数量不多。
        Me.txtSenderName.Text = dicCust.Item("Rst").Item(1).Item(1)
        Me.cboSenderProvince.Text = dicCust.Item("Rst").Item(1).Item(4)
        Me.txtSenderAddress.Text = dicCust.Item("Rst").Item(1).Item(5)
        Me.txtSenderPhone.Text = dicCust.Item("Rst").Item(1).Item(6)
        Me.txtSenderMobile.Text = dicCust.Item("Rst").Item(1).Item(7)
        Me.txtSenderCompany.Text = dicCust.Item("Rst").Item(1).Item(3)
    End If

End Sub

Private Sub Form_Load()
    bolisNew = True
    Call InitLayout(Me)
    Call InitTextBox(Me)
    Me.txtPaymentType.AddItem "月结"
    Me.txtPaymentType.AddItem "到付"
    
    Me.txtOrderID.Text = ""
    Me.txtCreateEmp.Text = gUSERNAME
    Set mControlsPOI = GetAllControlsPOI(Me)
    
    Call FillCboWithSampleDic(Me.cboReceiverProvince, gdicLocation.Item("0"))
    Call FillCboWithSampleDic(Me.cboSenderProvince, gdicLocation.Item("0"))
    
    'Call FillComboBoxWithDic(Me.txtSenderCity, gdicLocation.Item("0"), "1")
    
End Sub

Private Sub Form_Resize()
    Call ResizeFormControls(Me, mControlsPOI, True)
End Sub



Private Sub TmrSetSingleChkbox_Timer()

    If TypeName(Me.ActiveControl) = "CheckBox" Then
        'Debug.Print Me.ActiveControl.name
        Call MakeSingleCheck(Me, Me.ActiveControl)
    End If

End Sub

'======================================================================================================================
'======================================================================================================================
'======================================================================================================================
'======================================================================================================================
'======================================================================================================================
'
'Private Function MoveButton(ByRef btn As VB.TextBox, ByVal x As Single, ByVal Y As Single)
'
'    If dicLayout.Exists(btn.name) Then
'
'        btn.Move btn.Left + x - dicLayout.Item(btn.name).Item("OffsetX"), btn.Top + Y - dicLayout.Item(btn.name).Item("OffsetY")
'
'        DoEvents
'        dicLayout.Item(btn.name).Item("Top") = btn.Top
'        dicLayout.Item(btn.name).Item("Left") = btn.Left
'        dicLayout.Item(btn.name).Item("X") = 0
'        dicLayout.Item(btn.name).Item("Y") = 0
'        dicLayout.Item("txtReceiverName").Item("OffsetX") = 0
'        dicLayout.Item("txtReceiverName").Item("OffsetY") = 0
'    End If
'
'End Function
'
'
'
'
'Private Sub txtReceiverName_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'
'    If mbolLayout Then
'        If dicLayout.Exists("txtReceiverName") Then
'            Debug.Print x & ":::" & Y
'            dicLayout.Item("txtReceiverName").Item("X") = x
'            dicLayout.Item("txtReceiverName").Item("Y") = Y
'            dicLayout.Item("txtReceiverName").Item("OffsetX") = x - dicLayout.Item("txtReceiverName").Item("Left")
'            dicLayout.Item("txtReceiverName").Item("OffsetY") = Y - dicLayout.Item("txtReceiverName").Item("Top")
'
'        End If
'    End If
'
'End Sub
'
'Private Sub txtReceiverName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If mbolLayout Then
'    If Button = vbLeftButton Then
'        If dicLayout.Exists("txtReceiverName") Then
'            Debug.Print x & ":" & Y
'            Call MoveButton(Me.txtReceiverName, x - CSng(dicLayout.Item("txtReceiverName").Item("X")), Y - CSng(dicLayout.Item("txtReceiverName").Item("Y")))
'            'DoEvents
'        End If
'    End If
'End If
'End Sub
Private Sub txtCustCode_Change()

End Sub

Private Sub txtCustCode_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call cmdSearchCustCode_Click
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


Private Sub txtReceiverName_Change()
    
    Call ShowIncrementalSearchList("vwReceiver_Simple", "ReceiverName", "like", Me.txtReceiverName, Me.lstIncremental)
    

End Sub

Private Sub txtReceiverName_KeyDown(KeyCode As Integer, Shift As Integer)

    Call SelectIncrementalResult(KeyCode, Me.txtReceiverName, lstIncremental)
    
End Sub

Private Sub txtReceiverName_LostFocus()
    '根据客户的名称，搜索客户信息，填入下面的文本框中！
    Dim strCustName As String
    
    strCustName = Me.txtReceiverName.Text

    If strCustName <> "" Then
    
        Dim dicCust As Scripting.Dictionary
    
        Set dicCust = SearchCustInfoByName(strCustName, "vwReceiver_Simple", "ReceiverName")
        
        '没办法，控件名称无法映射，只能手工赋值，好在数量不多。
        Me.txtReceiverName.Text = dicCust.Item("Rst").Item(1).Item(1)
        Me.cboReceiverProvince.Text = dicCust.Item("Rst").Item(1).Item(3)
        Me.txtReceiverAddress.Text = dicCust.Item("Rst").Item(1).Item(4)
        Me.txtReceiverPhone.Text = dicCust.Item("Rst").Item(1).Item(5)
        Me.txtReceiverMobile.Text = dicCust.Item("Rst").Item(1).Item(6)
        Me.txtReceiverCompany.Text = dicCust.Item("Rst").Item(1).Item(2)
    End If

End Sub
