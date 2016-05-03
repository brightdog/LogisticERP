VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form frmOrder_Detail 
   AutoRedraw      =   -1  'True
   Caption         =   "订单详情"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14325
   Icon            =   "frmOrder_Detail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   14325
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Enabled         =   0   'False
      Height          =   435
      Left            =   9600
      TabIndex        =   60
      Top             =   9840
      Width           =   1275
   End
   Begin VB.CommandButton cmdAddExpressNO 
      Caption         =   "添加"
      Height          =   315
      Left            =   4440
      TabIndex        =   55
      Top             =   8340
      Width           =   975
   End
   Begin VB.ListBox lstThirdPartExpressNO 
      Height          =   1320
      Left            =   1320
      TabIndex        =   54
      Top             =   8760
      Width           =   2955
   End
   Begin VB.TextBox txtThirdPartExpressNO 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   52
      Top             =   8340
      Width           =   2235
   End
   Begin VB.ComboBox cboThirdPartCompany 
      Height          =   300
      Left            =   1320
      TabIndex        =   51
      Top             =   8340
      Width           =   675
   End
   Begin VB.CommandButton cmdShowRemark 
      Caption         =   "第三方物流信息"
      Height          =   315
      Left            =   4440
      TabIndex        =   50
      Top             =   8760
      Width           =   1695
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
      Left            =   9660
      TabIndex        =   41
      Text            =   "Order code"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.PictureBox picOrderID_Barcode 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      DrawWidth       =   2
      Height          =   750
      Left            =   7380
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   228
      TabIndex        =   49
      Top             =   480
      Width           =   3420
   End
   Begin VB.ListBox lstIncremental 
      Height          =   1680
      Left            =   180
      TabIndex        =   48
      Top             =   8460
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtReceiverMobile 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   4500
      TabIndex        =   10
      Text            =   "Receiver mobile"
      Top             =   4740
      Width           =   2055
   End
   Begin VB.TextBox txtSenderMobile 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   4500
      TabIndex        =   4
      Text            =   "Sender mobile"
      Top             =   2100
      Width           =   2055
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
      TabIndex        =   30
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
      Left            =   960
      TabIndex        =   12
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
      Left            =   960
      TabIndex        =   6
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   33
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
      TabIndex        =   32
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
      TabIndex        =   31
      Tag             =   "PaymentType"
      Top             =   4560
      Width           =   1635
   End
   Begin VB.CommandButton cmdSearchCustCode 
      Caption         =   "搜索"
      Height          =   315
      Left            =   6000
      TabIndex        =   2
      Top             =   1620
      Width           =   795
   End
   Begin VB.TextBox txtCustCode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   4080
      TabIndex        =   1
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
      TabIndex        =   19
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
      TabIndex        =   36
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      Left            =   12600
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
      Appearance      =   0
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
      Left            =   7500
      TabIndex        =   43
      Top             =   9840
      Width           =   1275
   End
   Begin VB.TextBox txtRemark 
      Height          =   1155
      Left            =   6840
      TabIndex        =   38
      Text            =   "备注："
      Top             =   8400
      Width           =   7455
   End
   Begin VB.TextBox txtReceiverSigner 
      Height          =   270
      Left            =   11940
      TabIndex        =   37
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
      Left            =   5700
      TabIndex        =   42
      Text            =   "Pickup receipt iD"
      Top             =   0
      Width           =   1635
   End
   Begin VB.TextBox txtOrderID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "Order iD"
      Top             =   1380
      Width           =   2295
   End
   Begin VB.CommandButton cmdSaveDetail 
      Caption         =   "保存"
      Height          =   435
      Left            =   4860
      TabIndex        =   39
      Top             =   9840
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
      TabIndex        =   34
      Text            =   "Other service"
      Top             =   5580
      Width           =   6315
   End
   Begin VB.TextBox txtReceiverAddress 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmOrder_Detail.frx":0856
      Top             =   5640
      Width           =   4395
   End
   Begin VB.TextBox txtReceiverPhone 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   960
      TabIndex        =   9
      Text            =   "Receiver phone"
      Top             =   4740
      Width           =   2055
   End
   Begin VB.TextBox txtReceiverCompany 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   1200
      TabIndex        =   11
      Text            =   "Receivercompany"
      Top             =   5160
      Width           =   5475
   End
   Begin VB.TextBox txtReceiverName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      TabIndex        =   8
      Text            =   "Receiver name"
      Top             =   4320
      Width           =   4275
   End
   Begin VB.TextBox txtSenderAddress 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmOrder_Detail.frx":0867
      Top             =   2880
      Width           =   4395
   End
   Begin VB.TextBox txtSenderPhone 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   960
      TabIndex        =   3
      Text            =   "Sender phone"
      Top             =   2100
      Width           =   2055
   End
   Begin VB.TextBox txtSenderCompany 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   1140
      TabIndex        =   5
      Text            =   "Sender company"
      Top             =   2520
      Width           =   5535
   End
   Begin VB.TextBox txtSenderName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      TabIndex        =   0
      Text            =   "Sender name"
      Top             =   1620
      Width           =   1695
   End
   Begin TDBNumber6Ctl.TDBNumber txtPkgLength 
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   7320
      Width           =   555
      _Version        =   65536
      _ExtentX        =   979
      _ExtentY        =   450
      Calculator      =   "frmOrder_Detail.frx":0876
      Caption         =   "frmOrder_Detail.frx":0896
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":08F9
      Keys            =   "frmOrder_Detail.frx":0917
      Spin            =   "frmOrder_Detail.frx":0961
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
      TabIndex        =   17
      Top             =   7320
      Width           =   555
      _Version        =   65536
      _ExtentX        =   979
      _ExtentY        =   450
      Calculator      =   "frmOrder_Detail.frx":0989
      Caption         =   "frmOrder_Detail.frx":09A9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":0A0C
      Keys            =   "frmOrder_Detail.frx":0A2A
      Spin            =   "frmOrder_Detail.frx":0A74
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
      TabIndex        =   18
      Top             =   7320
      Width           =   555
      _Version        =   65536
      _ExtentX        =   979
      _ExtentY        =   450
      Calculator      =   "frmOrder_Detail.frx":0A9C
      Caption         =   "frmOrder_Detail.frx":0ABC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":0B1F
      Keys            =   "frmOrder_Detail.frx":0B3D
      Spin            =   "frmOrder_Detail.frx":0B87
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
      TabIndex        =   35
      Top             =   7200
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   556
      Calendar        =   "frmOrder_Detail.frx":0BAF
      Caption         =   "frmOrder_Detail.frx":0CAA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOrder_Detail.frx":0D0D
      Keys            =   "frmOrder_Detail.frx":0D2B
      Spin            =   "frmOrder_Detail.frx":0D89
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
      ShowContextMenu =   -1
      ShowLiterals    =   2
      TabAction       =   0
      Text            =   ""
      ValidateMode    =   0
      ValueVT         =   2010185729
      Value           =   2.12482986761524E-314
      CenturyMode     =   0
   End
   Begin VB.Label Label4 
      Caption         =   "创建日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11340
      TabIndex        =   59
      Top             =   60
      Width           =   2115
   End
   Begin VB.Label Label3 
      Caption         =   "员工姓名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7680
      TabIndex        =   58
      Top             =   60
      Width           =   2115
   End
   Begin VB.Label Label2 
      Caption         =   "系统取件单号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4020
      TabIndex        =   57
      Top             =   60
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "系统订单编号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   56
      Top             =   60
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Label lblThirdPartCode 
      BackColor       =   &H0000FFFF&
      Caption         =   "第三方编号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   53
      Top             =   8400
      Width           =   1275
   End
   Begin VB.Image imgExpressNO_Barcode 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   2550
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
      TabIndex        =   47
      Top             =   1140
      Width           =   1515
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
      TabIndex        =   46
      Top             =   1140
      Width           =   1515
   End
   Begin VB.Image imgBackGround 
      Appearance      =   0  'Flat
      Height          =   7830
      Left            =   0
      Picture         =   "frmOrder_Detail.frx":0DB1
      Stretch         =   -1  'True
      Top             =   360
      Width           =   14265
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
Attribute VB_Name = "frmOrder_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dicLayout As Scripting.Dictionary
Dim mbolLayout As Boolean
Public mID As String
Private bolisNew As Boolean
Public mdicHiddenFields As New Scripting.Dictionary
Dim mControlsPOI As Scripting.Dictionary
Private mstrLastReceiverName As String

Dim dicOrderDetail As Scripting.Dictionary


Public Function LoadDetail(ByVal ID As String) As String
    bolisNew = False
    mID = ID
    Dim strURL, strPostData As String
    strPostData = "{""Type"":""LoadOrderDetailByID"",""Fields"":[""OrderID""],""Values"":[""" & ID & """]}"
    strURL = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strURL, strPostData)
    
    Set dicOrderDetail = JSON.Parse(strResult) '把结果存到窗体的字典对象中去，后续打印的时候，需要传到报表里去。
    
    Call FillFormTextBox(Me, dicOrderDetail)
    
    If Me.txtOrderID.Text <> "" Then
        Call Code39.Bar39(Me.picOrderID_Barcode, 5, Me.txtOrderID.Text, False, False)
        Me.cmdPrint.Enabled = True
    End If
    
    '    Me.imgExpressNO_Barcode.Picture = Me.picExpressNO_Barcode.Picture
    '    Me.imgExpressNO_Barcode.Refresh
    Me.lstIncremental.Visible = False
    '    If Me.txtThirdPartExpressNO.Text <> "" Then
    '        'Me.cmdSaveDetail.Enabled = False
    '        Call SetUILock(Me, False)
    '    End If
    
    If Me.txtRemark.Text = "" Then
        Me.txtRemark.Text = "备注："
    End If
    
    Call LoadThirdPartExpressNOList(ID)
    
End Function

Private Sub LoadThirdPartExpressNOList(ByVal ID As String)

    Dim strURL, strPostData As String
    strPostData = "{""Type"":""LoadlByID"",""Fields"":[""OrderID""],""Values"":[""" & ID & """]}"
    strURL = "loadthirdpartexpressnolist.asp"
    Dim strResult As String
    strResult = PostData(strURL, strPostData)
    'Call FillFormTextBox(Me, JSON.Parse(strResult))
    Dim dicList As Scripting.Dictionary
    Set dicList = JSON.Parse(strResult)
    Dim i As Integer
    Dim intExpressNOPos As Integer
    Dim intSiteCodePos As Integer
    
    For i = 1 To dicList.Item("Header").Count
    
        If dicList.Item("Header").Item(i) = "ExpressNO" Then
        
            intExpressNOPos = i
        ElseIf dicList.Item("Header").Item(i) = "SiteCode" Then
            intSiteCodePos = i
        End If
    
    Next
    Me.lstThirdPartExpressNO.Clear
    For i = 1 To dicList.Item("Rst").Count
    
        Call Me.lstThirdPartExpressNO.AddItem(dicList.Item("Rst").Item(i).Item(intSiteCodePos) & "_" & dicList.Item("Rst").Item(i).Item(intExpressNOPos))
    
    Next
    

End Sub

Public Function LoadDetailByExpressNO(ByVal ID As String) As String
    bolisNew = False
    mID = ID
    Dim strURL, strPostData As String
    strPostData = "{""Type"":""LoadOrderDetailByExpressNO"",""Fields"":[""ExpressNO""],""Values"":[""" & ID & """]}"
    strURL = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strURL, strPostData)
    Call FillFormTextBox(Me, JSON.Parse(strResult))
    Call Code39.Bar39(Me.picOrderID_Barcode, 5, Me.txtOrderID.Text, True, False)
    Me.lstIncremental.Visible = False
    If Me.txtThirdPartExpressNO.Text <> "" Then
        Me.cmdSaveDetail.Enabled = False
        Call SetUILock(Me, False)
    End If
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

Private Sub cmdAddExpressNO_Click()

    If Trim(Me.txtThirdPartExpressNO.Text) <> "" Then
        Me.lstThirdPartExpressNO.AddItem Me.cboThirdPartCompany.Text & "_" & Me.txtThirdPartExpressNO.Text
        Me.txtThirdPartExpressNO.Text = ""
    End If

End Sub

Private Sub cmdPrint_Click()
    Dim Rpt As rptOrderDetail
    Set Rpt = New rptOrderDetail
    
    Dim x As Variant
    
    For Each x In Rpt.PageHeader.Controls
    
        Select Case Left(x.name, 3)

            Case "txt", "cbo"
                x.Text = ""
                x.Text = CallByName(CallByName(Me, "controls", VbGet, x.name), "Text", VbGet)
    
            Case "chk"
                x.Caption = ""
                x.Value = CallByName(CallByName(Me, "controls", VbGet, x.name), "Value", VbGet)
                
            Case "lbl"
                x.Text = ""
                x.Text = CallByName(CallByName(Me, "controls", VbGet, x.name), "Caption", VbGet)
                
        End Select

    Next
    
    Rpt.picOrderID_Barcode.Picture = Me.picOrderID_Barcode.Picture
    'Call Rpt.InitReport(dicOrderDetail)
    
    Rpt.Show vbModal
    Set Rpt = Nothing
End Sub

Private Sub cmdSaveDetail_Click()

    If bolisNew Then
        Me.txtCreateDT.Text = Format(VBA.Now(), "yyyy-mm-dd hh:mm:ss")
    End If
    
    Me.txtThirdPartExpressNO.Text = Trim(Replace(Replace(Me.txtThirdPartExpressNO.Text, vbCr, ""), vbLf, ""))
    
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
    
    If Me.lstThirdPartExpressNO.ListCount > 0 Then '为一对多父子关系的订单号和快递单号做关联，需要单独处理
    
        SBField.Append """ThirdPartExpressNOList"","
        Dim i As Integer
        Dim strThirdPartExpressNOList As String

        For i = 0 To Me.lstThirdPartExpressNO.ListCount - 1
        
            strThirdPartExpressNOList = strThirdPartExpressNOList & Me.lstThirdPartExpressNO.List(i) & "|"
        
        Next

        SBValue.Append """" & strThirdPartExpressNOList & ""","
    End If
    
    Dim strPostData As String
    Dim strURL As String
    Dim strFields As String
    Dim strValues As String
    strFields = SBField.toString
    strValues = SBValue.toString

    strPostData = "{""Type"":""SaveOrderDetail"",""Fields"":[" & Left(strFields, Len(strFields) - 1) & "],""Values"":[" & Left(strValues, Len(strValues) - 1) & "]}"
    strURL = LCase(Me.name) & ".asp"
    Dim strResult As String
    strResult = PostData(strURL, strPostData)
    Debug.Print strResult
    
    Dim dicResult As Scripting.Dictionary
    
    Set dicResult = JSON.Parse(strResult)

    If Not dicResult Is Nothing Then
        If dicResult.Item("STATE") <> "ERR" Then '流氓判断法，先把信息提示出来再说了。
            If MsgBox("保存已成功，是否打印？", vbYesNo + vbQuestion, "操作提示") = vbYes Then
                Call LoadDetail(GetOrderIDFromDic(dicResult))
                Me.cmdPrint.Enabled = True
                Call cmdPrint_Click
            Else
                Unload Me
            End If

        Else
            MsgBox "保存失败，请检查字段信息"
            
            '        If Me.txtThirdPartExpressNO.Text <> "" Then
            '
            '            Call VBA.Shell(App.path & "\ExpressBot.exe " & Me.cboThirdPartCompany.Text & "_" & Me.txtThirdPartExpressNO, vbHide)
            '
            '        Else
            '
            '        End If
        
        End If

    Else
        WriteLog "####" & strURL & vbCrLf & "####" & strPostData & vbCrLf & "####" & strResult
        MsgBox "其他错误，请联系系统管理员"
    
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
        '<EhHeader>
        On Error GoTo cmdSearchCustCode_Click_Err
        '</EhHeader>
        Dim strCustCode As String
    
100     strCustCode = Me.txtCustCode.Text

102     If strCustCode <> "" Then
    
            Dim dicCust As Scripting.Dictionary
    
104         Set dicCust = SearchCustInfoByCode(strCustCode)
        
            '没办法，控件名称无法映射，只能手工赋值，好在数量不多。
106         If dicCust.Item("Rst").Count > 0 Then
108             Me.txtSenderName.Text = dicCust.Item("Rst").Item(1).Item(1)
110             Me.cboSenderProvince.Text = dicCust.Item("Rst").Item(1).Item(7)
112             Me.txtSenderAddress.Text = dicCust.Item("Rst").Item(1).Item(4)
114             Me.txtSenderPhone.Text = dicCust.Item("Rst").Item(1).Item(5)
116             Me.txtSenderMobile.Text = dicCust.Item("Rst").Item(1).Item(6)
118             Me.txtSenderCompany.Text = dicCust.Item("Rst").Item(1).Item(3)
            Else
            
                 MsgBox "当前客户编号不存在，请重新输入。"
                 Me.txtCustCode.SetFocus
                
            End If
        End If

        '<EhFooter>
        Exit Sub

cmdSearchCustCode_Click_Err:
        WriteLog Err.Description & vbCrLf & _
               "in LogisticERP.frmOrder_Detail.cmdSearchCustCode_Click " & _
               "at line " & Erl
        MsgBox "当前客户编号不存在，请重新输入！"
        Me.txtCustCode.SetFocus
        
        '</EhFooter>
End Sub

Private Sub cmdShowRemark_Click()

    Dim strExpressNOList As String
    Dim i As Integer
    
    For i = 0 To Me.lstThirdPartExpressNO.ListCount - 1

        If Me.lstThirdPartExpressNO.List(i) <> "" Then
            strExpressNOList = strExpressNOList & Me.lstThirdPartExpressNO.List(i) & "|"
        End If

    Next
    Call ShowExpressDetailWindow(strExpressNOList)
End Sub

Private Sub ShowExpressDetailWindow(ByVal strExpressNOList As String)

    Dim objExpressDetail As frmVenderExpressDetail
    Set objExpressDetail = New frmVenderExpressDetail
    Call objExpressDetail.LoadDetail(strExpressNOList)
    objExpressDetail.Show vbModal
    Set objExpressDetail = Nothing

End Sub

Private Sub Form_Load()
    bolisNew = True
    Call InitLayout(Me)
    Call InitTextBox(Me)
    'Me.txtPaymentType.AddItem "月结"
    'Me.txtPaymentType.AddItem "到付"
    
    Me.txtOrderID.Text = ""
    Me.txtCreateEmp.Text = gUSERNAME
    Set mControlsPOI = GetAllControlsPOI(Me)
    
    Call FillCboWithSampleDic(Me.cboReceiverProvince, gdicLocation.Item("0"))
    Call FillCboWithSampleDic(Me.cboSenderProvince, gdicLocation.Item("0"))
    Me.cboThirdPartCompany.AddItem "ZTO"
    Me.cboThirdPartCompany.ListIndex = 0
    
    'Call FillComboBoxWithDic(Me.txtSenderCity, gdicLocation.Item("0"), "1")
    
End Sub

Private Sub Form_Resize()
    Call ResizeFormControls(Me, mControlsPOI, True)
End Sub







Private Sub lstThirdPartExpressNO_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Call MouseClick(0, 0)

        DoEvents

        If Me.lstThirdPartExpressNO.List(Me.lstThirdPartExpressNO.ListIndex) <> "" Then
            Debug.Print (x + Me.lstThirdPartExpressNO.Left + Me.Left) / Screen.TwipsPerPixelX & ":" & (y + Me.lstThirdPartExpressNO.Height + Me.Height) / Screen.TwipsPerPixelY
            Me.PopupMenu mnuOper
        End If
    End If
End Sub

Private Sub mnuDel_Click(Index As Integer)
    If MsgBox("删除当前纪录？点订单保存按钮之后生效！", vbOKCancel) = vbOK Then
    
        Debug.Print "执行删除操作"
        Call Me.lstThirdPartExpressNO.RemoveItem(Me.lstThirdPartExpressNO.ListIndex)
    End If
End Sub

Private Sub mnuOpenDetail_Click(Index As Integer)
    Call ShowExpressDetailWindow(Me.lstThirdPartExpressNO.Text)
End Sub

Private Sub TmrSetSingleChkbox_Timer()

    If TypeName(Me.ActiveControl) = "CheckBox" Then
        Debug.Print Me.ActiveControl.name
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
    Call ShowIncrementalSearchList("tblCust", "CustCode", "like", Me.txtCustCode, Me.lstIncremental)
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


Private Sub txtCustCode_LostFocus()
    Call lstIncremental_DBlClick
End Sub

Private Sub txtReceiverName_Change()
    
    Call ShowIncrementalSearchList("vwReceiver_Simple", "ReceiverName", "like", Me.txtReceiverName, Me.lstIncremental)
    

End Sub

Private Sub txtReceiverName_GotFocus()
    mstrLastReceiverName = Me.txtReceiverName.Text
End Sub

Private Sub txtReceiverName_KeyDown(KeyCode As Integer, Shift As Integer)

    Call SelectIncrementalResult(KeyCode, Me.txtReceiverName, lstIncremental)
    
End Sub

Private Sub txtReceiverName_LostFocus()
    '根据客户的名称，搜索客户信息，填入下面的文本框中！
    Dim strCustName As String
    
    strCustName = Me.txtReceiverName.Text
    If mstrLastReceiverName = strCustName Then
        Exit Sub
    End If
    If strCustName <> "" Then
    
        Dim dicCust As Scripting.Dictionary
    
        Set dicCust = SearchCustInfoByName(strCustName, "vwReceiver_Simple", "ReceiverName")
        
        '没办法，控件名称无法映射，只能手工赋值，好在数量不多。
        If dicCust.Item("Rst").Count > 0 Then
            Me.txtReceiverName.Text = dicCust.Item("Rst").Item(1).Item(1)
            Me.cboReceiverProvince.Text = dicCust.Item("Rst").Item(1).Item(6)
            Me.txtReceiverAddress.Text = dicCust.Item("Rst").Item(1).Item(3)
            Me.txtReceiverPhone.Text = dicCust.Item("Rst").Item(1).Item(4)
            Me.txtReceiverMobile.Text = dicCust.Item("Rst").Item(1).Item(5)
            Me.txtReceiverCompany.Text = dicCust.Item("Rst").Item(1).Item(2)
        End If
    End If

End Sub

Private Sub txtRemark_Click()
    If Me.txtRemark.Text = "备注：" Then
        Me.txtRemark.Text = ""
    End If
End Sub

Private Sub txtRemark_LostFocus()
    If Me.txtRemark.Text = "" Then
        Me.txtRemark.Text = "备注："
    End If
End Sub
