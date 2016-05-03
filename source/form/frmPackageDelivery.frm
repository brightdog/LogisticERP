VERSION 5.00
Begin VB.Form frmPackageDelivery 
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   Icon            =   "frmPackageDelivery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9285
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印出库单"
      Height          =   435
      Left            =   7380
      TabIndex        =   6
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox txtExpressNo 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Text            =   "Express no"
      Top             =   720
      Width           =   4335
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
      Height          =   4620
      Left            =   4620
      TabIndex        =   5
      Top             =   1320
      Width           =   4515
   End
   Begin VB.TextBox txtPackageDeliveryReceiptID 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Text            =   "Out warehouse receipt no"
      Top             =   0
      Width           =   4335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "加入"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   3
      Top             =   720
      Width           =   1395
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   1
      Top             =   0
      Width           =   1395
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   435
      Left            =   4020
      TabIndex        =   4
      Top             =   6120
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "快递单号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   9
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "当前派件单编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   8
      Top             =   60
      Width           =   3135
   End
   Begin VB.Label lblPackageDeliveryReceiptInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   4515
   End
End
Attribute VB_Name = "frmPackageDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mControlsPOI As Scripting.Dictionary

Private Sub cmdAdd_Click()
    
    Dim strPackageDeliveryReceiptNO As String
    Dim strExpressNO As String
    
    Dim strJSON As String
    
    strJSON = "{"
    strJSON = strJSON & """PackageDeliveryReceiptID"":""" & Me.txtPackageDeliveryReceiptID.Text & """"
    strJSON = strJSON & ",""ExpressNO"":""" & Me.txtExpressNo.Text & """"
    '    strJSON = strJSON & ",""WareHouseName"":""" & Me.cboWareHouseName.Text & """"
    strJSON = strJSON & "}"
    Dim strResult As String
    strResult = modPostData_Core.PostData("savesingleexpressnotoPackageDeliveryreceipt.asp", strJSON)
    Debug.Print strResult
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)
    
    Select Case CStr(dicResult.Item("ERR"))
        
        Case "InventorySTATEERR"
            MsgBox "运单中转状态错误"
            Exit Sub

        Case "STATEERR"
            MsgBox "其他错误"
            Exit Sub
    
    End Select
    
    If CStr(dicResult.Item("ERR")) = "" Then
        Me.lstExpressNo.AddItem Me.txtExpressNo.Text
        Me.lstExpressNo.Selected(Me.lstExpressNo.ListCount - 1) = True
    Else
    
        Dim i As Integer
        
        For i = 0 To Me.lstExpressNo.ListCount - 1
        
            If Me.lstExpressNo.List(i) = Me.txtExpressNo.Text Then
                Me.lstExpressNo.Selected(i) = True
                Exit For
            End If
        
        Next
    
    End If

    Me.txtExpressNo.Text = ""
End Sub

Private Sub cmdClose_Click()
    
    Call frmMain.CloseTab(Me.Tag)
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Call Search(Me.txtPackageDeliveryReceiptID.Text)
End Sub

Public Function Search(ByVal PackageDeliveryReceiptID As String) As Boolean

    If Trim(PackageDeliveryReceiptID) <> "" Then
        Debug.Print "txtPackageDeliveryReceiptID_LostFocus:" & PackageDeliveryReceiptID

        If Trim(Me.txtPackageDeliveryReceiptID.Text) <> "" Then
            Debug.Print "txtPackageDeliveryReceiptID_LostFocus:" & Me.txtPackageDeliveryReceiptID.Text
            Dim dicPackageDeliveryReceiptInfo As Scripting.Dictionary
            Set dicPackageDeliveryReceiptInfo = CheckPackageDeliveryReceiptInfoByID(Me.txtPackageDeliveryReceiptID.Text)

            If dicPackageDeliveryReceiptInfo.Item("ERR") = Empty Then
            
                Dim strInfo As String
                strInfo = ""
                Dim i As Integer
            
                For i = 1 To dicPackageDeliveryReceiptInfo.Item("Header").Count
                    strInfo = strInfo & dicPackageDeliveryReceiptInfo.Item("Header")(i) & ":" & vbTab & dicPackageDeliveryReceiptInfo.Item("Rst")(1)(i) & vbCrLf

                    '                    If dicPackageDeliveryReceiptInfo.Item("Header")(i) = "WarehouseID" Then
                    '
                    '                        Dim dicWarehouse As Scripting.Dictionary
                    '                        Set dicWarehouse = GetWarehouseListbyID(dicPackageDeliveryReceiptInfo.Item("Rst")(1)(i))
                    '
                    '                        Call FillComboBoxWithDic(Me.cboWareHouseName, dicWarehouse, "WarehouseName")
                    '
                    '                    End If

                Next
            
                Me.lblPackageDeliveryReceiptInfo.Caption = strInfo
            
                Dim dicList As Scripting.Dictionary
        
                Set dicList = loadCurrentPackageDeliveryReceiptsExpressNoList(Me.txtPackageDeliveryReceiptID.Text)
        
                Call FillListBoxWithDic(Me.lstExpressNo, dicList, "ExpressNO")
        
                Me.cmdAdd.Enabled = True
        
            Else
                'Me.cboWareHouseName.Enabled = False
                Me.cmdAdd.Enabled = False
                Me.lstExpressNo.Clear
    
            End If
        End If
    End If

End Function

Private Sub Form_Load()
    Set mControlsPOI = GetAllControlsPOI(Me)
    Me.txtExpressNo.Text = ""
    Me.txtPackageDeliveryReceiptID.Text = ""
End Sub

Private Sub lstExpressNo_DblClick()

    If Me.lstExpressNo.ListIndex >= 0 Then
        If RemoveExpressNOFromPackageDeliveryReceipt(Me.txtPackageDeliveryReceiptID.Text, Me.lstExpressNo.List(Me.lstExpressNo.ListIndex)) Then
        
            Me.lstExpressNo.RemoveItem Me.lstExpressNo.ListIndex
        Else
            MsgBox "操作失败！"
        End If
    End If

End Sub

Private Function RemoveExpressNOFromPackageDeliveryReceipt(ByVal PackageDeliveryReceiptID As String, ByVal ExpressNO As String) As Boolean
    
    Dim strPickupReceiptNO As String
    Dim strExpressNO As String
    
    Dim strJSON As String
    
    strJSON = "{""PackageDeliveryReceiptID"":""" & PackageDeliveryReceiptID & """,""ExpressNO"":""" & ExpressNO & """}"
    Dim strResult As String
    strResult = modPostData_Core.PostData("removesingleexpressnofrompackagedeliveryreceipt.asp", strJSON)
    Debug.Print strResult
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)

    If dicResult.Item("State") = "SUCCESS" Then
        
        RemoveExpressNOFromPackageDeliveryReceipt = True
    Else
        RemoveExpressNOFromPackageDeliveryReceipt = False
    
    End If
    
End Function
