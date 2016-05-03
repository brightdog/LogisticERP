VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "条形码生成系统-----枕善居收集整理"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4935
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2640
      List            =   "Form1.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "预览"
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4755
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   735
         Left            =   120
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   93
         TabIndex        =   4
         Top             =   360
         Width           =   1392
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "750103131130"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "复制到剪切板"
      Height          =   390
      Left            =   2760
      TabIndex        =   0
      Top             =   3525
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   855
      Picture         =   "Form1.frx":0036
      Top             =   4170
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "类型："
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "信息："
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/****************************************************************************
' * Summary   : 条形码生成程序
' * Version   : 1.00
' * Start Date: 2004-6-07
' * My home   : http://www.mndsoft.com
' * E-Mail    : Mnd@Mndsoft.Com
' ****************************************************************************/

Dim cl As New arisBarcode
'Dim xArr As New XArrayDB

Private Sub Command1_Click()
    Select Case Combo1.ListIndex
    Case 0
         cl.Code128 Picture1, 6, Text1, True
    Case 1
         cl.Code39 Picture1, 6, Text1, True
    Case 2
         cl.EAN13 Picture1, 6, Text1, True
    Case 3
         cl.EAN8 Picture1, 6, Text1, True
    End Select
End Sub

Private Sub Command2_Click()
    Clipboard.Clear
    Clipboard.SetData Picture1.Image, 2
End Sub


