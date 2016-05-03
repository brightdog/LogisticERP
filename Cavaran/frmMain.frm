VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E5B0E85C-65F0-11D2-ACBA-0080ADA85544}#1.0#0"; "JSBBar16.ocx"
Begin VB.Form frmMain 
   Caption         =   "LogisticERP_Caravan"
   ClientHeight    =   5970
   ClientLeft      =   1950
   ClientTop       =   2685
   ClientWidth     =   9645
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   9645
   Begin VB.PictureBox picPagging 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4140
      ScaleHeight     =   285
      ScaleWidth      =   5445
      TabIndex        =   10
      Top             =   5640
      Width           =   5475
      Begin VB.CommandButton cmdPaggingLast 
         Height          =   315
         Index           =   10
         Left            =   4860
         TabIndex        =   24
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdPaggingNext 
         Height          =   315
         Index           =   10
         Left            =   4500
         TabIndex        =   23
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdPaggingPrev 
         Height          =   315
         Index           =   10
         Left            =   480
         TabIndex        =   22
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdPaggingFirst 
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdPagging 
         Height          =   315
         Index           =   9
         Left            =   4140
         TabIndex        =   20
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         Height          =   315
         Index           =   8
         Left            =   3780
         TabIndex        =   19
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         Height          =   315
         Index           =   7
         Left            =   3420
         TabIndex        =   18
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         Height          =   315
         Index           =   6
         Left            =   3060
         TabIndex        =   17
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         Height          =   315
         Index           =   5
         Left            =   2700
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         Height          =   315
         Index           =   4
         Left            =   2340
         TabIndex        =   15
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         Height          =   315
         Index           =   3
         Left            =   1980
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         Height          =   315
         Index           =   2
         Left            =   1620
         TabIndex        =   13
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPagging 
         Height          =   315
         Index           =   0
         Left            =   900
         TabIndex        =   11
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   4320
      ScaleHeight     =   2565
      ScaleWidth      =   4245
      TabIndex        =   6
      Top             =   1440
      Width           =   4245
      Begin VB.Image imgLogo 
         Height          =   2085
         Left            =   0
         Picture         =   "frmMain.frx":030A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3120
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   741
      ButtonWidth     =   609
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "iml16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1200
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Edit"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PrintPreview"
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Summary"
            Object.ToolTipText     =   "View Summary..."
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Group"
            Object.ToolTipText     =   "Group by..."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sort"
            Object.ToolTipText     =   "Sort..."
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowGroupByBox"
            Object.ToolTipText     =   "Group By Box"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "AllowAddNew"
            Object.ToolTipText     =   "Show Allow Add New Row"
            ImageIndex      =   10
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MoveFirst"
            Object.ToolTipText     =   "Move first visible record"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MovePrevious"
            Object.ToolTipText     =   "Move previous visible record"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MoveNext"
            Object.ToolTipText     =   "Move next visible record"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MoveLast"
            Object.ToolTipText     =   "Move last visible record"
            ImageIndex      =   16
         EndProperty
      EndProperty
      Begin MSComDlg.CommonDialog cdlNWind 
         Left            =   8925
         Top             =   15
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Where is JSNWind.mdb?"
         Filter          =   "Databases|*.mdb"
      End
      Begin VB.CommandButton cmdNewRecord 
         Appearance      =   0  'Flat
         Caption         =   "New Record"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   30
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   30
         Width           =   1260
      End
      Begin VB.ComboBox cboStyle 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4545
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   45
         Width           =   1965
      End
   End
   Begin VB.PictureBox picTvw 
      BorderStyle     =   0  'None
      Height          =   4110
      Left            =   1455
      ScaleHeight     =   274
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   146
      TabIndex        =   8
      Top             =   1200
      Width           =   2190
      Begin MSComctlLib.TreeView tvwCatalog 
         Height          =   4950
         Left            =   165
         TabIndex        =   9
         Top             =   120
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   8731
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   265
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "iml16"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin GridEX20.GridEX grDropDown 
      Height          =   2520
      Index           =   0
      Left            =   5865
      TabIndex        =   7
      Top             =   1785
      Visible         =   0   'False
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   4445
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   "CustomerID"
      ReplaceColumnIndex=   "Company Name"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      HeaderStyle     =   3
      ContScroll      =   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmMain.frx":13B00
      DataMode        =   1
      GridLines       =   0
      ColumnHeaderHeight=   285
      ColumnsCount    =   3
      Column(1)       =   "frmMain.frx":13E1A
      Column(2)       =   "frmMain.frx":13F72
      Column(3)       =   "frmMain.frx":140AE
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmMain.frx":141F2
      FormatStyle(2)  =   "frmMain.frx":1432A
      FormatStyle(3)  =   "frmMain.frx":143DA
      FormatStyle(4)  =   "frmMain.frx":1448E
      FormatStyle(5)  =   "frmMain.frx":14566
      ImageCount      =   1
      ImagePicture(1) =   "frmMain.frx":1461E
      PrinterProperties=   "frmMain.frx":14938
   End
   Begin GridEX20.GridEX jsgxMain 
      Height          =   4530
      Left            =   3720
      TabIndex        =   4
      Top             =   1020
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7990
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      GridLineStyle   =   2
      DetectRowDrag   =   -1  'True
      MultiSelect     =   -1  'True
      HideSelection   =   1
      Options         =   8
      RecordsetType   =   1
      ForeColorHeader =   0
      AllowDelete     =   -1  'True
      BorderStyle     =   2
      MaskColor       =   16711935
      ImageCount      =   1
      ImagePicture1   =   "frmMain.frx":14B10
      DataMode        =   1
      HeaderFontName  =   "Tahoma"
      HeaderFontBold  =   -1  'True
      HeaderFontWeight=   700
      FontName        =   "Tahoma"
      CardWidth       =   0
      AllowAddNew     =   -1  'True
      ColumnHeaderHeight=   285
      ColumnsCount    =   14
      Column(1)       =   "frmMain.frx":14E2A
      Column(2)       =   "frmMain.frx":14F9A
      Column(3)       =   "frmMain.frx":151E6
      Column(4)       =   "frmMain.frx":15402
      Column(5)       =   "frmMain.frx":155E6
      Column(6)       =   "frmMain.frx":1580A
      Column(7)       =   "frmMain.frx":15A2E
      Column(8)       =   "frmMain.frx":15CBE
      Column(9)       =   "frmMain.frx":15F12
      Column(10)      =   "frmMain.frx":16166
      Column(11)      =   "frmMain.frx":163B2
      Column(12)      =   "frmMain.frx":16616
      Column(13)      =   "frmMain.frx":168A6
      Column(14)      =   "frmMain.frx":16AD2
      GroupCount      =   1
      Group(1)        =   "frmMain.frx":16D62
      FmtConditionsCount=   3
      ApplyGroupCondition=   -1  'True
      GroupConditionCountTitle=   "On Sale"
      GroupCondition  =   "frmMain.frx":16DCA
      FmtCondition(1) =   "frmMain.frx":16EA6
      FmtCondition(2) =   "frmMain.frx":16FCA
      FmtCondition(3) =   "frmMain.frx":170C2
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmMain.frx":171B6
      FormatStyle(2)  =   "frmMain.frx":172BA
      FormatStyle(3)  =   "frmMain.frx":1736A
      FormatStyle(4)  =   "frmMain.frx":1741E
      FormatStyle(5)  =   "frmMain.frx":174F6
      ImageCount      =   1
      ImagePicture(1) =   "frmMain.frx":175AE
      PrinterProperties=   "frmMain.frx":178C8
   End
   Begin JSBtnBar16.ButtonBar jsbbMain 
      Align           =   3  'Align Left
      Height          =   5550
      Left            =   0
      Top             =   420
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   9790
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GroupsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      ButtonStyle     =   1
      ButtonGroupStyle=   -1  'True
      LargeImagesCount=   5
      LargeImageHeight=   32
      LargeImageWidth =   32
      LargeImagePicture1=   "frmMain.frx":17AA0
      LargeImagePicture2=   "frmMain.frx":17DBA
      LargeImagePicture3=   "frmMain.frx":180D4
      LargeImagePicture4=   "frmMain.frx":183EE
      LargeImagePicture5=   "frmMain.frx":18708
      SmallImagesCount=   5
      SmallImageHeight=   16
      SmallImageWidth =   16
      SmallImagePicture1=   "frmMain.frx":18A22
      SmallImagePicture2=   "frmMain.frx":18D3C
      SmallImagePicture3=   "frmMain.frx":19056
      SmallImagePicture4=   "frmMain.frx":19370
      SmallImagePicture5=   "frmMain.frx":1968A
      GroupCount      =   2
      GroupCaption1   =   "Catalogs"
      GroupToolTipText1=   "NorthWind Traders catalogs"
      GroupCaption2   =   "Orders"
      GroupToolTipText2=   "Northwind Traders Orders"
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   7590
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":199A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A2F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A60C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A926
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AC40
            Key             =   "Group"
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AF5A
            Key             =   "Sort"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B274
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B58E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B8A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BBC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BEDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C1F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C510
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C82A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CB44
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CC56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   8175
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CD68
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D082
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D39C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D6B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D9D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DCEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblfront 
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BackStyle       =   0  'Transparent
      Caption         =   "Caravan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   435
      Left            =   1605
      TabIndex        =   1
      Top             =   480
      Width           =   1305
   End
   Begin VB.Image imgCat 
      Height          =   480
      Left            =   7035
      Picture         =   "frmMain.frx":1E004
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblback 
      BackColor       =   &H80000010&
      Height          =   555
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   7530
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewRecord 
         Caption         =   ""
      End
      Begin VB.Menu mnuEditRecord 
         Caption         =   ""
      End
      Begin VB.Menu mnuSepRecord 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnuPrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
         Index           =   0
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Selected Items..."
         Index           =   1
      End
      Begin VB.Menu mnuSepPrint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuCurrentView 
         Caption         =   "CurrentView"
         Begin VB.Menu MnuViewStyle 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuViewSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnushow 
         Caption         =   "Show Fields..."
      End
      Begin VB.Menu mnusort 
         Caption         =   "Sort..."
      End
      Begin VB.Menu mnugroup 
         Caption         =   "Group By..."
      End
      Begin VB.Menu mnuformat 
         Caption         =   "Format View..."
      End
      Begin VB.Menu mnuEXCol 
         Caption         =   "Expand/Collapse Groups"
         Begin VB.Menu mnucolall 
            Caption         =   "Collapse All"
         End
         Begin VB.Menu mnuexpall 
            Caption         =   "Expand All"
         End
      End
      Begin VB.Menu mnuSum 
         Caption         =   "View Summary..."
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuButtonBar 
         Caption         =   "Janus Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFoldersList 
         Caption         =   "Folders List"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnugbBox 
         Caption         =   "Group By Box"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Advanced Sample"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim m_CurrentView As Long

Dim m_CatalogIndex As Long
Dim m_conn As Connection
Dim m_LastBaseIcon As Integer
Dim mvarBookmark As Variant
Dim mcolForms As New Collection
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub imgCat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        Width = Width - 15
    End If
End Sub

Public Sub OnRecordUpdate(ByVal CatalogIndex As Integer, Bookmark As Variant)

    
    If m_CatalogIndex = CatalogIndex Then
        jsgxMain.Refresh
        If IsNull(Bookmark) Then
            jsgxMain.SearchNewRecords
        Else
            jsgxMain.MoveToBookmark Bookmark
            jsgxMain.Update
        End If
    End If
End Sub
Private Sub cboStyle_Click()
Dim rstViews As Recordset
Dim rsViewDetails As Recordset
Dim rsCatalogDetails As Recordset
Dim col As JSColumn
Dim lngColumnId As Long
Dim IsTableView As Boolean
Dim prevVisble As Boolean
Dim T As Long
    T = timeGetTime
    If cboStyle.ListIndex = -1 Then
        m_CurrentView = 0
        Exit Sub
    End If
    Screen.MousePointer = 11
    SaveCurrentView
    prevVisble = jsgxMain.Visible
    jsgxMain.Visible = False
    m_CurrentView = cboStyle.ItemData(cboStyle.ListIndex)
    Set rstViews = New Recordset
    rstViews.Open "SELECT * FROM Views WHERE ViewID=" & m_CurrentView, m_conn, adOpenForwardOnly, adLockReadOnly
    If Not rstViews.EOF Then
        If Not IsNull(rstViews![Layout]) Then jsgxMain.LoadLayoutString rstViews![Layout], False
    End If
    LoadValueLists
    IsTableView = (jsgxMain.View = jgexTable)
    tlbMain.Buttons("ShowGroupByBox").Enabled = IsTableView
    If jsgxMain.GroupByBoxVisible Then
        tlbMain.Buttons("ShowGroupByBox").Value = tbrPressed
    Else
        tlbMain.Buttons("ShowGroupByBox").Value = tbrUnpressed
    End If
    If jsgxMain.AllowAddNew Then
        tlbMain.Buttons("AllowAddNew").Value = tbrPressed
    Else
        tlbMain.Buttons("AllowAddNew").Value = tbrUnpressed
    End If
    tlbMain.Buttons("Group").Enabled = IsTableView
    tlbMain.Buttons("AllowAddNew").Enabled = IsTableView
    mnugbBox.Enabled = IsTableView
    mnugroup.Visible = IsTableView
    mnuEXCol.Visible = IsTableView
    CheckMnuViewStyle
    Screen.MousePointer = 0
    jsgxMain.Visible = prevVisble

    On Error Resume Next
    jsgxMain.SetFocus
    T = timeGetTime - T
    Debug.Print "ChangeView", T / 1000
End Sub



Private Sub cmdNewRecord_Click()

    ShowRecord True
    
End Sub

Private Sub Form_Load()

    Form_Resize
    m_LastBaseIcon = jsgxMain.GridImages.Count
    If LoadCatalogSettings Then
        SetCatalog
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim i As Long
Dim LeftTemp As Long
    If Me.Width < 15000 Then
        Me.Width = 15000
    End If
    If Me.Height < 10000 Then
        Me.Height = 10000
    End If
    jsbbMain.Visible = mnuButtonBar.Checked
    picTvw.Visible = mnuFoldersList.Checked
    picTvw.Top = lblback.Top + lblback.Height + 60
    If mnuButtonBar.Checked Then
        LeftTemp = jsbbMain.Width
    End If
    lblback.Move LeftTemp, lblback.Top, ScaleWidth - LeftTemp, lblback.Height
    LeftTemp = LeftTemp + 60
    picTvw.Left = LeftTemp
    lblfront.Move LeftTemp, lblfront.Top ', ScaleWidth - LeftTemp
    If mnuFoldersList.Checked Then
        LeftTemp = picTvw.Left + picTvw.Width + 60
    End If
    i = ScaleWidth - 650
    If i < lblfront.Left + lblfront.Width + 60 Then i = lblfront.Left + lblfront.Width + 60
    imgCat.Move i, imgCat.Top, imgCat.Width, imgCat.Height
    picTvw.Height = ScaleHeight - picTvw.Top - 60
    'i = pictvw.left + 60 + pictvw.width
    jsgxMain.Move LeftTemp, picTvw.Top, ScaleWidth - LeftTemp - 30, ScaleHeight - picTvw.Top - 60
    With jsgxMain
        picMain.Move .Left, .Top, .Width, .Height
        imgLogo.Move 0, 0, .Width, .Height
        imgLogo.Refresh
'        imgLogo.Top = 0
'        imgLogo.Left = 0
'        imgLogo.Width = picMain.Width
'        imgLogo.Height = picMain.Height
    End With
    
     
     
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim frm As Object
On Error GoTo EH_Unload
    SaveCurrentView
    For Each frm In mcolForms
        Unload frm
    Next
    Exit Sub

EH_Unload:
    MsgBox Err.Description
    
End Sub

Private Sub imgCat_DblClick()

    Me.Width = Me.Width + 15
    
End Sub

Private Sub jsbbMain_ItemClick(ByVal Item As JSBtnBar16.JSGroupItem)
Dim nod As Node

    Set nod = tvwCatalog.Nodes(Item.Key)
    nod.Selected = True
    tvwCatalog_NodeClick nod
    
End Sub





Private Sub jsgxMain_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)

    jsgxMain.PrinterProperties.FooterString(jgexHFRight) = "Page " & PageNumber & " of " & nPages
    
End Sub

Private Sub jsgxMain_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
Dim rstCatDetails As Recordset
Dim grTemp As JSGroup
On Error GoTo EH_ColumnHeaderClick

    If Column.IsGrouped Then
        For Each grTemp In jsgxMain.Groups
            If grTemp.ColIndex = Column.Index Then
                jsgxMain_GroupByBoxHeaderClick grTemp
            End If
        Next
    Else
        If Column.SortOrder = 0 Then
            jsgxMain.SortKeys.Clear
            jsgxMain.SortKeys.Add Column.Index, jgexSortAscending
        Else
            If Column.SortOrder = 1 Then
                jsgxMain.SortKeys.Clear
                jsgxMain.SortKeys.Add Column.Index, -1
            Else
                jsgxMain.SortKeys.Clear
                jsgxMain.SortKeys.Add Column.Index, 1
            End If
        End If
    End If
    Exit Sub
    
EH_ColumnHeaderClick:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub jsgxMain_DblClick()
    
    ShowRecord
End Sub



Private Sub jsgxMain_GroupByBoxHeaderClick(ByVal Group As GridEX20.JSGroup)

    Screen.MousePointer = 11
    Group.SortOrder = -Group.SortOrder
    jsgxMain.RefreshGroups
    Screen.MousePointer = 0
End Sub

Private Sub jsgxMain_OLESetData(Data As GridEX20.JSDataObject, DataFormat As Integer)

    Data.SetData jsgxMain.GetClipString(True), jgexCFText
    
End Sub

Private Sub jsgxMain_OLEStartDrag(Data As GridEX20.JSDataObject, AllowedEffects As Long)

    AllowedEffects = jgexDropEffectCopy
    Data.SetData , jgexCFText
    
End Sub

Private Sub jsgxMain_RowDrag(ByVal Button As Integer, ByVal Shift As Integer)

    Debug.Print "RowDrag"
    jsgxMain.ExpandSelection
    jsgxMain.OLEDrag
    
End Sub

Private Sub jsgxMain_RowFormat(RowBuffer As GridEX20.JSRowData)
Dim strGroupCaption As String

    If RowBuffer.RowType = jgexRowTypeGroupHeader Then
        'Adding the count of records in a group for group headers
        strGroupCaption = RowBuffer.GroupCaption & " (" & RowBuffer.RecordCount & " " & lblfront.Caption & ") "
        RowBuffer.GroupCaption = strGroupCaption
    End If
End Sub

Private Sub mnuAbout_Click()

    frmAbout.Show 1
    
End Sub

Private Sub mnuButtonBar_Click()

    mnuButtonBar.Checked = Not mnuButtonBar.Checked
    Form_Resize
    
End Sub

Private Sub mnucolall_Click()
    jsgxMain.CollapseAll
End Sub

Private Sub mnuEditRecord_Click()
    ShowRecord
End Sub

Private Sub mnuExit_Click()

    Unload Me
    
End Sub


Private Sub mnuexpall_Click()
    jsgxMain.ExpandAll
End Sub

Private Sub mnuFile_Click()
Dim strCatalog As String

    If m_CatalogIndex <> 0 Then
        strCatalog = Left(lblfront.Caption, Len(lblfront.Caption) - 1)
        mnuNewRecord.Caption = "New " & strCatalog
        mnuEditRecord.Caption = " Edit " & strCatalog
    End If
    mnuPrintPreview.Enabled = (m_CatalogIndex <> 0)
    mnuPrint(0).Enabled = (m_CatalogIndex <> 0)
    mnuPrint(1).Enabled = (m_CatalogIndex <> 0)


End Sub

Private Sub mnuFoldersList_Click()

    mnuFoldersList.Checked = Not mnuFoldersList.Checked
    Form_Resize
    
End Sub

Private Sub mnuformat_Click()
    If Not jsgxMain.Visible Then Exit Sub
    If jsgxMain.View = jgexTable Then
        frmTableview.FormatGrid jsgxMain
    Else
        frmCardView.FormatGrid jsgxMain
    End If
    
End Sub

Private Sub mnugbBox_Click()

    mnugbBox.Checked = Not mnugbBox.Checked
    GroupByBoxVisible mnugbBox.Checked
End Sub

Private Sub mnugroup_Click()
    frmGroupBy.GroupGrid jsgxMain
'    CheckGroups
End Sub

Private Sub mnuNewRecord_Click()
    ShowRecord True
End Sub

Private Sub mnuPageSetup_Click()
    jsgxMain.PrinterProperties.PageSetup Me.hWnd
End Sub

Private Sub mnuPrint_Click(Index As Integer)
On Error GoTo EH_mnuPrint
    With jsgxMain.PrinterProperties
        .TranslateColors = True
        .RepeatHeaders = True
        .HeaderString(jgexHFCenter) = "Janus GridEX Advanced Sample" & vbCrLf & lblfront.Caption
        .FooterString(jgexHFLeft) = Format(Date, "Medium Date")
    End With
    If Index = 0 Then
        jsgxMain.PrintGrid True
    Else
        jsgxMain.PrintGrid True, True
    End If
    Exit Sub
EH_mnuPrint:
    MsgBox Err.Description, vbExclamation
    
End Sub

Private Sub mnuPrintPreview_Click()
On Error GoTo EH_mnuPrintPreview
    With jsgxMain.PrinterProperties
        .TranslateColors = True
        .RepeatHeaders = True
        .HeaderString(jgexHFCenter) = "Janus GridEX Advanced Sample" & vbCrLf & lblfront.Caption
        .FooterString(jgexHFLeft) = Format(Date, "Medium Date")
    End With
    Load frmPrintPreview
    If Me.WindowState = vbMaximized Then
        frmPrintPreview.WindowState = vbMaximized
    Else
        frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    End If
    frmPrintPreview.Show
    DoEvents
    Screen.MousePointer = 11
    jsgxMain.PrintPreview frmPrintPreview.GEXPreview1, (jsgxMain.SelectedItems.Count > 1)
    Screen.MousePointer = 0
    Hide
    Exit Sub
EH_mnuPrintPreview:
    MsgBox Err.Description, vbExclamation
    Unload frmPrintPreview
    Screen.MousePointer = 0
    
End Sub

Private Sub mnushow_Click()
    
    If Not jsgxMain.Visible Then Exit Sub
    frmShowfields.ShowFields jsgxMain
   
End Sub

Private Sub mnusort_Click()
    If Not jsgxMain.Visible Then Exit Sub
    frmSort.SortGrid jsgxMain
End Sub

Private Sub mnuSum_Click()
    frmSummary.ShowSummary jsgxMain
End Sub




Private Sub MnuViewStyle_Click(Index As Integer)
    
    cboStyle.ListIndex = Index
End Sub

Private Sub picMain_Paint()

    With picMain
        .Cls
        picMain.Line (0, .ScaleHeight)-(0, 0), vbButtonShadow
        picMain.Line (0, 0)-(.ScaleWidth, 0), vbButtonShadow
        picMain.Line (.ScaleWidth - 1, 0)-(.ScaleWidth - 1, .ScaleHeight), vb3DHighlight
        picMain.Line (.ScaleWidth, .ScaleHeight - 1)-(0, .ScaleHeight - 1), vb3DHighlight
    End With
End Sub

Private Sub picMain_Resize()
Dim lTemp As Long
Dim tTemp As Long

    picMain_Paint
    lTemp = (picMain.ScaleWidth - imgLogo.Width) \ 2
    tTemp = (picMain.ScaleHeight - imgLogo.Height) \ 2
    If lTemp < 1 Then lTemp = 1
    If tTemp < 1 Then tTemp = 1
    imgLogo.Move lTemp, tTemp
    
    
End Sub


Private Sub picTvw_Paint()

    With picTvw
        picTvw.Line (0, .ScaleHeight)-(0, 0), vbButtonShadow
        picTvw.Line (0, 0)-(.ScaleWidth, 0), vbButtonShadow
        picTvw.Line (.ScaleWidth - 1, 0)-(.ScaleWidth - 1, .ScaleHeight), vb3DHighlight
        picTvw.Line (.ScaleWidth, .ScaleHeight - 1)-(0, .ScaleHeight - 1), vb3DHighlight
    End With
End Sub

Private Sub picTvw_Resize()
On Error Resume Next
    picTvw_Paint
    tvwCatalog.Move 1, 1, picTvw.ScaleWidth - 1, picTvw.ScaleHeight
    
End Sub


Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintPreview"
            mnuPrintPreview_Click
        Case "Print"
            mnuPrint_Click 0
        Case "Edit"
            ShowRecord
        Case "Group"
            If Not jsgxMain.Visible Then Exit Sub
            frmGroupBy.GroupGrid jsgxMain
        Case "Sort"
            If Not jsgxMain.Visible Then Exit Sub
            frmSort.SortGrid jsgxMain
        Case "ShowGroupByBox"
            jsgxMain.GroupByBoxVisible = (Button.Value = tbrPressed)
            mnugbBox.Checked = (Button.Value = tbrPressed)
        Case "AllowAddNew"
            jsgxMain.AllowAddNew = (Button.Value = tbrPressed)
        Case "Summary"
            frmSummary.ShowSummary jsgxMain
        Case "MoveFirst"
            jsgxMain.MoveFirst
            jsgxMain.SetFocus
        Case "MoveLast"
            jsgxMain.MoveLast
            jsgxMain.SetFocus
        Case "MovePrevious"
            jsgxMain.MovePrevious
            jsgxMain.SetFocus
        Case "MoveNext"
            jsgxMain.MoveNext
            jsgxMain.SetFocus
    End Select
End Sub

Private Sub tvwCatalog_Click()

    tvwCatalog_NodeClick tvwCatalog.SelectedItem
    
End Sub

Private Function LoadCatalogSettings() As Boolean
    Dim rstCatalogs As Recordset
    Dim nod As Node
    Dim lngParentIndex As Long
    Dim i As Integer
    Dim dbName As String
    Dim ConnString As String
    Dim Key As String
    Dim GroupItems As JSGroupItems
    Dim bbItem As JSGroupItem
    On Error GoTo EH_LoadCatalog
    
'100     dbName = App.Path & "\JSNWind.MDB"
'102     dbName = GetSetting("Janus Advanced Sample", "Initial Settings", "DBPath", dbName)
        dbName = "Cavaran"
OpenDatabase_Proc:
104     ConnString = "Provider=SQLOLEDB.1;Persist Security InFso=true;Data Source='114.215.177.126,55944';User ID='sa';Password='Sdfg2345';Initial Catalog='" & dbName & "'"
106     Set m_conn = New Connection
108     m_conn.Open ConnString
110     SaveSetting "Janus Advanced Sample", "Initial Settings", "DBPath", dbName
112     jsgxMain.DatabaseName = ConnString
114     Set rstCatalogs = New Recordset
116     rstCatalogs.Open "SELECT * FROM Catalogs", m_conn, adOpenStatic, adLockOptimistic
118     m_CatalogIndex = -1
120     Set nod = tvwCatalog.Nodes.Add(, , , "Cavaran", 1)
122     nod.Expanded = True
124     nod.Tag = 0
126     lngParentIndex = nod.Index
128     i = 0
130     Do Until rstCatalogs.EOF
132         i = i + 1
134         Key = "K" & rstCatalogs![CatalogId]
136         Set nod = tvwCatalog.Nodes.Add(lngParentIndex, tvwChild, Key, rstCatalogs![Name], CInt(rstCatalogs![IconIndex]))
138         nod.Tag = rstCatalogs![CatalogId]
140         If i = 5 Then
142             Set GroupItems = jsbbMain.ButtonGroups(2).GroupItems
            Else
144             Set GroupItems = jsbbMain.ButtonGroups(1).GroupItems
            End If
146         Set bbItem = GroupItems.Add(, Key, nod.Text, i, i)
148         bbItem.ToolTipText = "Cavaran's " & bbItem.Caption
150         rstCatalogs.MoveNext
        Loop
152     tvwCatalog.Nodes(lngParentIndex).Selected = True
154     LoadCatalogSettings = True
        Exit Function
    
EH_LoadCatalog:
156     Select Case Err.Number
            Case &H80004005 'Couldn't Find File
                Dim strTemp As String
           
158             strTemp = GetFileName(dbName)
160             If strTemp = "" Then
                    Exit Function
                Else
162                 dbName = strTemp
164                 Resume OpenDatabase_Proc
                End If
166         Case Else
168             MsgBox Err.Description, vbCritical
        End Select
        
End Function

Private Function GetFileName(strFileName As String) As String
On Error GoTo EH_GetFileName

    cdlNWind.CancelError = True
    cdlNWind.Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames
    cdlNWind.FileName = strFileName
    cdlNWind.InitDir = App.Path
    cdlNWind.ShowOpen
    GetFileName = cdlNWind.FileName

    
    Exit Function
    
EH_GetFileName:
    MsgBox "Operation Canceled", vbCritical
    
End Function


Private Sub SetCatalog()
Dim Node As Node
Dim i As Integer
Dim c As JSColumn
Dim T As Long

    T = timeGetTime
    Screen.MousePointer = 11

    Set Node = tvwCatalog.SelectedItem
    m_CatalogIndex = CLng(Node.Tag)
    lblfront.Caption = Node.Text
    jsgxMain.Visible = False
    Me.mnuNewRecord.Visible = (m_CatalogIndex <> 0)
    Me.mnuEditRecord.Visible = (m_CatalogIndex <> 0)
    Me.mnuSepRecord.Visible = (m_CatalogIndex <> 0)
    If m_CatalogIndex = 0 Then
        mnuViewEnabled False
        EnableButtons False
        picMain.Visible = True
        cboStyle.Clear
        cboStyle.Enabled = False
        jsgxMain.Visible = False
    Else
        SaveCurrentView
        m_CurrentView = 0
        mnuViewEnabled True
        EnableButtons True
        picMain.Visible = False
        LoadCatalog
        cboStyle.Enabled = True
        jsgxMain.Visible = True
        jsgxMain.Row = 1
    End If
    T = timeGetTime - T
    Debug.Print "SetCatalog", T / 1000
    Screen.MousePointer = 0
 
    
End Sub


Private Sub GroupByBoxVisible(ByVal Value As Boolean)
Dim b As MSComctlLib.Button

    Set b = tlbMain.Buttons("ShowGroupByBox")
    If (b.Value = tbrPressed) = Value Then Exit Sub
    If Value Then
        b.Value = tbrPressed
    Else
        b.Value = tbrUnpressed
    End If
    tlbMain_ButtonClick b
    
    
End Sub

Public Sub AllowAddNew(ByVal Value As Boolean)
Dim b As MSComctlLib.Button

    Set b = tlbMain.Buttons("AllowAddNew")
    If (b.Value = tbrPressed) = Value Then Exit Sub
    If Value Then
        b.Value = tbrPressed
    Else
        b.Value = tbrUnpressed
    End If
    tlbMain_ButtonClick b
    
    
End Sub
Private Sub LoadColumns()
'Dim rstCatDetails As Recordset
'Dim rstValueList As Recordset
'Dim IconIndex As Integer
'Dim col As JSColumn
'Dim picTemp As IPictureDisp
'On Error Resume Next
'    jsgxMain.Columns.Clear
'    Set rstCatDetails = m_db.OpenRecordset("CatalogDetails")
'    rstCatDetails.Index = "CatalogIndex"
'    rstCatDetails.Seek ">=", m_CatalogIndex
'    If Not rstCatDetails.NoMatch Then
'        Do Until rstCatDetails![CatalogId] <> m_CatalogIndex
'            If IsNull(rstCatDetails![DataField]) Then
'                Set col = jsgxMain.Columns.Add()
'            Else
'                Set col = jsgxMain.Columns.Add(, , , rstCatDetails![DataField])
'            End If
'            With col
'                .Caption = rstCatDetails![Caption]
'                .ColumnType = rstCatDetails![ColumnType]
'                .EditType = rstCatDetails![EditType]
'                .CardCaption = rstCatDetails![IsCardTitle]
'                .CardIcon = rstCatDetails![IscardIcon]
'                .DataField = rstCatDetails![DataField]
'                .DefaultIcon = rstCatDetails![DefaultIcon]
'                .FetchData = rstCatDetails![FetchData]
'                .FetchIcon = rstCatDetails![FetchIcon]
'                .Format = rstCatDetails![Format]
'                .GroupEmptyStringCaption = rstCatDetails![GroupEmptyStringCaption]
'                .GroupFormat = rstCatDetails![GroupFormat]
'                .GroupPrefix = rstCatDetails![GroupPrefix]
'                .HasValueList = rstCatDetails![HasValueList]
'                .SortType = rstCatDetails![SortType]
'                .TextAlignment = rstCatDetails![TextAlignment]
'                .Width = rstCatDetails![Width]
'                .AllowSizing = rstCatDetails![AllowSizing]
'                .SortType = rstCatDetails![SortType]
'                .Tag = rstCatDetails![Description]
'                .Visible = True
'                If .HasValueList Then
'                    Set rstValueList = m_db.OpenRecordset(rstCatDetails![ValueListRecordSource], dbOpenSnapshot)
'                    Do Until rstValueList.EOF
'                        IconIndex = 0
'                        If Not IsNull(rstValueList![PictureFile]) Then
'                            Set picTemp = LoadPicture(App.Path & "\Icons\" & rstValueList![PictureFile])
'                            If Not picTemp Is Nothing Then
'                                jsgxMain.GridImages.Add picTemp
'                                IconIndex = jsgxMain.GridImages.Count
'                            End If
'                        End If
'                        .ValueList.Add rstValueList![Value], rstValueList![Text], IconIndex
'                        rstValueList.MoveNext
'                    Loop
'                    Set rstValueList = Nothing
'                End If
'            End With
'            rstCatDetails.MoveNext
'            If rstCatDetails.EOF Then Exit Do
'        Loop
'    End If
End Sub

Private Sub LoadCatalog(Optional ByVal intPage As Integer = 1)
Dim rstCatalog As Recordset
Dim DefaultView As Variant
Dim i As Long
On Error Resume Next
    Set rstCatalog = New Recordset
    rstCatalog.Open "SELECT * FROM Catalogs WHERE CatalogID=" & m_CatalogIndex, m_conn, adOpenForwardOnly, adLockReadOnly
    If Not rstCatalog.EOF Then
        Dim strSP As String
        strSP = rstCatalog.Fields.Item("RecordSource").Value
        Set rstCatalog = m_conn.Execute("EXEC " & strSP & " " & intPage & " " & gListPageSize)
        With jsgxMain
            .ClearFields
            'Call FillGrid(jsgxMain, rstCatalog)
            '.RecordSource =
            .Recordset = rstCatalog
            .Rebind
            LoadViews
            DefaultView = rstCatalog![DefaultView]
            If IsNull(DefaultView) Then
                cboStyle.ListIndex = 0
            Else
                cboStyle.ListIndex = DefaultView
            End If
        End With
    End If
    
End Sub

Private Sub LoadViews()
Dim rstViews As Recordset
Dim i As Integer
On Error Resume Next
    Set rstViews = New Recordset
    rstViews.Open "SELECT * FROM Views WHERE CatalogID=" & m_CatalogIndex, m_conn, adOpenForwardOnly, adLockReadOnly
    For i = 1 To cboStyle.ListCount - 1
        Unload MnuViewStyle(i)
    Next
    cboStyle.Clear
    Do Until rstViews.EOF
        cboStyle.AddItem rstViews![Name]
        cboStyle.ItemData(cboStyle.NewIndex) = rstViews![ViewId]
        rstViews.MoveNext
        If rstViews.EOF Then Exit Do
    Loop
    MnuViewStyle(0).Caption = cboStyle.List(0)
    For i = 1 To cboStyle.ListCount - 1
        Load MnuViewStyle(i)
        MnuViewStyle(i).Caption = cboStyle.List(i)
        MnuViewStyle(i).Visible = True
    Next
End Sub



Public Sub UnloadForm(FormKey As String)

    mcolForms.Remove FormKey
    
End Sub

Private Sub tvwCatalog_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim tmpIndex As Long
    Dim BarItem As JSGroupItem
    Dim T As Long
    T = timeGetTime
    
        
    If Node Is Nothing Then Exit Sub

    Set Node = tvwCatalog.SelectedItem
    tmpIndex = CLng(Node.Tag)
    If tmpIndex = m_CatalogIndex Then Exit Sub
    Set imgCat.Picture = iml32.ListImages(Node.Image).ExtractIcon
    m_CatalogIndex = tmpIndex
    SetCatalog
    Select Case m_CatalogIndex
        Case 0
            Set jsbbMain.SelectedItem = Nothing
        Case 5
            Set BarItem = jsbbMain.ButtonGroups(2).GroupItems(1)
            BarItem.Selected = True
        Case Else
            Set BarItem = jsbbMain.ButtonGroups(1).GroupItems(m_CatalogIndex)
            BarItem.Selected = True
    End Select
    T = timeGetTime - T
    Debug.Print "Time = " & T / 1000 & " secs."
    
    
End Sub



Private Sub EnableButtons(ByVal bEnable As Boolean)
Dim b As Button

    For Each b In tlbMain.Buttons
        If b.Style = tbrDefault Or b.Style = tbrCheck Then
            b.Enabled = bEnable
        End If
    Next
    cmdNewRecord.Enabled = bEnable
    
End Sub

Private Function GetKeyFromBookmark(strPrefix As String, Bookmark As Variant) As String
Dim strTemp As String
Static NullCount As Long
Dim i As Integer
    If IsNull(Bookmark) Then
        NullCount = NullCount + 1
        strTemp = NullCount
    Else
        strTemp = strTemp & Bookmark
    End If
    GetKeyFromBookmark = strPrefix & strTemp


End Function

Private Sub ShowRecord(Optional NewRecord As Boolean)
Dim strKey As String
Dim varBookmark As Variant
Dim RowIndex As Long
Dim frmTemp As Form
Dim rs As Recordset
    If m_CatalogIndex = 0 Then Exit Sub
    strKey = "Cat" & m_CatalogIndex & "-"
    If NewRecord Then
        varBookmark = Null
    Else
        RowIndex = jsgxMain.RowIndex(jsgxMain.Row)
        If RowIndex = 0 Then
            Exit Sub
        Else
            varBookmark = jsgxMain.RowBookmark(RowIndex)
        End If
    End If
    strKey = GetKeyFromBookmark(strKey, varBookmark)
    On Error Resume Next
    Set frmTemp = mcolForms.Item(strKey)
    If Err Then
        Select Case m_CatalogIndex
            Case CatalogCustomers
                Set frmTemp = New frmCustomers
            Case CatalogSuppliers
                Set frmTemp = New frmSuppliers
            Case CatalogEmployees
                Set frmTemp = New frmEmployees
            Case CatalogProducts
                Set frmTemp = New frmProducts
            Case CatalogOrders
                Set frmTemp = New frmOrders
        End Select
        frmTemp.Key = strKey
        mcolForms.Add frmTemp, strKey
        Set rs = jsgxMain.ADORecordset
        If IsNull(varBookmark) Then
            frmTemp.NewRecord m_conn, rs
        Else
            rs.Bookmark = varBookmark
            frmTemp.EditRecord m_conn, rs
        End If
    Else
        If frmTemp.WindowState = vbMinimized Then
            frmTemp.WindowState = vbNormal
        End If
        frmTemp.SetFocus
    End If

End Sub



Private Sub CheckMnuViewStyle()
Dim mnu As Menu

    For Each mnu In MnuViewStyle
        mnu.Checked = (mnu.Index = cboStyle.ListIndex)
    Next
End Sub

Private Sub mnuViewEnabled(ByVal Enabled As Boolean)

    mnuCurrentView.Enabled = Enabled
    mnushow.Enabled = Enabled
    mnusort.Enabled = Enabled
    mnugroup.Enabled = Enabled
    mnuformat.Enabled = Enabled
    mnuEXCol.Enabled = Enabled
    mnuSum.Enabled = Enabled
    mnugbBox.Enabled = Enabled
    
End Sub

Private Sub SaveCurrentView()
On Error GoTo EH_SaveCurrentView
Dim i As Long
Dim c As JSColumn
Dim rs As Recordset
Dim btArray() As Byte
    If m_CurrentView = 0 Then Exit Sub
    For i = jsgxMain.GridImages.Count To 3 Step -1
        jsgxMain.GridImages.Remove i
    Next
    For Each c In jsgxMain.Columns
        If c.HasValueList Then
            c.ValueList.Clear
        End If
        If Not c.DropDownControl Is Nothing Then
            Set c.DropDownControl = Nothing
        End If
    Next
    Set rs = New Recordset
    rs.Open "SELECT * FROM Views WHERE ViewID=" & m_CurrentView, m_conn, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        btArray = jsgxMain.LayoutString(True)
        rs![Layout] = btArray()
        rs.Update
    End If
    Exit Sub
    
EH_SaveCurrentView:
    MsgBox Err.Description, vbExclamation
    
End Sub

Private Sub LoadValueLists()
        '<EhHeader>
        On Error GoTo LoadValueLists_Err
        '</EhHeader>
    On Error GoTo EH_LoadValueList
    Dim rs As Recordset
    Dim c As JSColumn
    Dim vl As JSValueList
    Dim rsSource As Recordset
    Dim IconIndex As Long
    Dim Pic As StdPicture
    Dim GridsCount As Long
    Dim i As Long

100     Set rs = New Recordset
102     rs.Open "SELECT * FROM ColumnDetails WHERE CatalogID=" & m_CatalogIndex, m_conn, adOpenForwardOnly, adLockReadOnly
104     Do Until rs.EOF
106         Set c = jsgxMain.Columns(rs![ColumnKey])
108         If rs!HasValueList Then
110             c.HasValueList = True
112             Set vl = c.ValueList
114             vl.Clear
116             Set rsSource = New Recordset
118             rsSource.Open (rs![RecordSource]), m_conn
120             Do Until rsSource.EOF
122                 If Not IsNull(rsSource![PictureFile]) Then
124                     Set Pic = LoadPicture(App.Path & "\Icons\" & rsSource![PictureFile])
126                     IconIndex = jsgxMain.GridImages.Add(Pic).Index
                    Else
128                     IconIndex = 0
                    End If
130                 vl.Add rsSource(0), rsSource(1), IconIndex
132                 rsSource.MoveNext
                Loop
            
            Else
    '            Exit Sub
134             GridsCount = GridsCount + 1
136             If grDropDown.Count < GridsCount Then
138                 Load grDropDown(GridsCount - 1)
                End If
140             Set rsSource = New Recordset
142             rsSource.Open (rs![RecordSource]), m_conn, adOpenStatic, adLockReadOnly
144             With grDropDown(GridsCount - 1)
146                 .HoldFields
148                 Set .ADORecordset = rsSource
150                 .LoadLayoutString rs![Layout], False
152                 .Height = .ColumnHeaderHeight + .RowHeight * 8 + 2 * Screen.TwipsPerPixelY
                End With
154             Set c.DropDownControl = grDropDown(GridsCount - 1)
            End If
156         rs.MoveNext
        Loop
        Exit Sub
    
EH_LoadValueList:
158     MsgBox Err.Description & " " & Err.Number, vbExclamation
        '<EhFooter>
        Exit Sub

LoadValueLists_Err:
        MsgBox Err.Description & vbCrLf & _
               "in Cavaran.frmMain.LoadValueLists " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
