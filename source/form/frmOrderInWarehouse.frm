VERSION 5.00
Begin VB.Form frmOrderInWarehouse 
   AutoRedraw      =   -1  'True
   ClientHeight    =   7455
   ClientLeft      =   210
   ClientTop       =   210
   ClientWidth     =   9405
   ControlBox      =   0   'False
   Icon            =   "frmOrderInWarehouse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   9405
   StartUpPosition =   1  '����������
   Tag             =   "������"
   Begin VB.ComboBox cboWareHouseName 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3660
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   4335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�"
      Height          =   435
      Left            =   3600
      TabIndex        =   10
      Top             =   6900
      Width           =   1995
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "��ѯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8100
      TabIndex        =   1
      Top             =   240
      Width           =   1275
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8100
      TabIndex        =   4
      Top             =   1440
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ��ⵥ"
      Height          =   435
      Left            =   7080
      TabIndex        =   8
      Top             =   6900
      Width           =   2235
   End
   Begin VB.TextBox txtExpressNo 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3660
      TabIndex        =   3
      Text            =   "Express no"
      Top             =   1440
      Width           =   4335
   End
   Begin VB.ListBox lstExpressNo 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      Left            =   4800
      TabIndex        =   6
      Top             =   2040
      Width           =   4515
   End
   Begin VB.TextBox txtOutWarehouseReceiptNo 
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3660
      TabIndex        =   0
      Text            =   "Pickup receipt no"
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "��ǰ�ռ��ֿ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   180
      TabIndex        =   11
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblInWarehouseReceiptInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   180
      TabIndex        =   9
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "��ݵ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   180
      TabIndex        =   7
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "��ǰ��ת���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   180
      TabIndex        =   5
      Top             =   300
      Width           =   3135
   End
   Begin VB.Menu mnuOper 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenDetail 
         Caption         =   "������"
         Index           =   1
      End
      Begin VB.Menu mnuDel 
         Caption         =   "ɾ��"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmOrderInWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mControlsPOI As Scripting.Dictionary

Private Sub cboWareHouseName_Click()

    If Me.cboWareHouseName.ListIndex > -1 Then
        Me.cmdAdd.Enabled = True
    End If

End Sub

Private Sub cmdAdd_Click()
    
    Dim strOutWarehouseReceiptNO As String
    Dim strExpressNO As String
    
    Dim strJSON As String
    
    strJSON = "{"
    strJSON = strJSON & """OutWarehouseExpressID"":""" & Me.txtOutWarehouseReceiptNo.Text & """"
    strJSON = strJSON & ",""ExpressNO"":""" & Me.txtExpressNo.Text & """"
    strJSON = strJSON & ",""WareHouseName"":""" & Me.cboWareHouseName.Text & """"
    strJSON = strJSON & "}"
    Dim strResult As String
    strResult = modPostData_Core.PostData("savesingleexpressnotoinwarehousereceipt.asp", strJSON)
    Debug.Print strResult
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)
    
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

        Select Case CStr(dicResult.Item("ERR"))
        
            Case "InventorySTATEERR"
                MsgBox "�˵���ת״̬����"
                Exit Sub

            Case "EIDNOTEXIST"
                MsgBox "�˵��Ų�����"
                Exit Sub

            Case Else
                MsgBox "��������"
                Exit Sub
        End Select

    End If

    Me.txtExpressNo.Text = ""
End Sub

Private Sub cmdClose_Click()
    
    Call frmMain.CloseTab(Me.Tag)
    Unload Me
End Sub

Public Function Search(ByVal OutWarehouseReceiptNo As String) As Boolean

    If Trim(OutWarehouseReceiptNo) <> "" Then
        Debug.Print "txtOutWarehouseReceiptNo_LostFocus:" & OutWarehouseReceiptNo

        If Trim(Me.txtOutWarehouseReceiptNo.Text) <> "" Then
            Debug.Print "txtOutWarehouseReceiptNo_LostFocus:" & Me.txtOutWarehouseReceiptNo.Text
            Dim dicOutWarehouseReceiptInfo As Scripting.Dictionary
            Set dicOutWarehouseReceiptInfo = CheckOutWarehouseReceiptInfoByID(Me.txtOutWarehouseReceiptNo.Text)

            If dicOutWarehouseReceiptInfo.Item("ERR") = Empty Then
            
                Dim strInfo As String
                strInfo = ""
                Dim i As Integer
                Dim strToWarehouseID As String
                
                For i = 1 To dicOutWarehouseReceiptInfo.Item("Header").Count
                    strInfo = strInfo & dicOutWarehouseReceiptInfo.Item("Header")(i) & ":" & vbTab & dicOutWarehouseReceiptInfo.Item("Rst")(1)(i) & vbCrLf

                    If dicOutWarehouseReceiptInfo.Item("Header")(i) = "WarehouseID" Then
                        strToWarehouseID = dicOutWarehouseReceiptInfo.Item("Rst")(1)(i)
                        Dim dicWarehouse As Scripting.Dictionary
                        Set dicWarehouse = GetWarehouseListbyID(strToWarehouseID)
                    
                        Call FillComboBoxWithDic(Me.cboWareHouseName, dicWarehouse, "WarehouseName")
                
                    End If

                Next
            
                Me.lblInWarehouseReceiptInfo.Caption = strInfo
            
                Dim dicList As Scripting.Dictionary
        
                Set dicList = loadCurrentInWarehouseReceiptsExpressNoList(Me.txtOutWarehouseReceiptNo.Text)
        
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

Private Sub cmdSearch_Click()

    Call Search(Me.txtOutWarehouseReceiptNo.Text)

End Sub

Private Sub Form_Load()
    Set mControlsPOI = GetAllControlsPOI(Me)
    Me.txtExpressNo.Text = ""
    Me.txtOutWarehouseReceiptNo.Text = ""

End Sub

Private Sub Form_Resize()
    Call ResizeFormControls(Me, mControlsPOI, True)
End Sub



Private Sub lstExpressNo_DblClick()

    If Me.lstExpressNo.ListIndex >= 0 Then
        Load frmOrder_Detail
        Call frmOrder_Detail.LoadDetailByExpressNO(Me.lstExpressNo.List(Me.lstExpressNo.ListIndex))
        frmOrder_Detail.Show vbModal
        Unload frmOrder_Detail
    End If

End Sub

Private Sub RemoveFromList(ListIndex)

    If Me.lstExpressNo.ListIndex >= 0 Then
        If RemoveExpressNOFromInWarehouseReceipt(Me.txtOutWarehouseReceiptNo.Text, Me.lstExpressNo.List(Me.lstExpressNo.ListIndex)) Then
        
            Me.lstExpressNo.RemoveItem Me.lstExpressNo.ListIndex
        Else
            MsgBox "����ʧ�ܣ�"
        End If
    End If

End Sub

Private Function RemoveExpressNOFromInWarehouseReceipt(ByVal OutWarehouseReceiptNo As String, ByVal ExpressNO As String) As Boolean
    
    Dim strPickupReceiptNO As String
    Dim strExpressNO As String
    
    Dim strJSON As String
    
    strJSON = "{""OutWarehouseReceiptNo"":""" & OutWarehouseReceiptNo & """,""ExpressNO"":""" & ExpressNO & """}"
    Dim strResult As String
    strResult = modPostData_Core.PostData("removesingleexpressnofrominwarehousereceipt.asp", strJSON)
    Debug.Print strResult
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)

    If dicResult.Item("State") = "SUCCESS" Then
        
        RemoveExpressNOFromInWarehouseReceipt = True
    Else
        RemoveExpressNOFromInWarehouseReceipt = False
    
    End If
    
End Function

Private Sub lstExpressNo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    If Button = vbRightButton Then
        Call MouseClick(0, 0)

        DoEvents

        If Me.lstExpressNo.List(Me.lstExpressNo.ListIndex) <> "" Then
            Debug.Print (x + Me.lstExpressNo.Left + Me.Left) / Screen.TwipsPerPixelX & ":" & (y + Me.lstExpressNo.Height + Me.Height) / Screen.TwipsPerPixelY
            Me.PopupMenu mnuOper
        End If
    End If

End Sub

Private Sub mnuDel_Click(Index As Integer)
    If MsgBox("ȷ��ɾ����ǰ��¼��", vbOKCancel) = vbOK Then
    
        Debug.Print "ִ��ɾ������"
        Call RemoveFromList(Me.lstExpressNo.ListIndex)
    End If
End Sub

Private Sub mnuOpenDetail_Click(Index As Integer)
    Call lstExpressNo_DblClick
End Sub

Private Sub txtExpressNo_GotFocus()
    Me.txtExpressNo.SelStart = 0
    Me.txtExpressNo.SelLength = Len(Me.txtExpressNo.Text)
End Sub

Private Sub txtExpressNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call cmdAdd_Click
    End If

End Sub

Private Sub txtOutWarehouseReceiptNo_GotFocus()
    Me.txtOutWarehouseReceiptNo.SelStart = 0
    Me.txtOutWarehouseReceiptNo.SelLength = Len(Me.txtOutWarehouseReceiptNo.Text)
    Debug.Print "txtOutWarehouseReceiptNo_GotFocus"
End Sub

Private Sub txtOutWarehouseReceiptNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call cmdSearch_Click
    End If

End Sub

Private Sub txtOutWarehouseReceiptNo_LostFocus()

    '�ȼ�⵱ǰȡ���������Ƿ���Ч
    '�����Ч���������ڵ�ǰȡ�����У������µĶ������
    '���Ƿ�Ҫ��һ���Ƴ������Ĺ����أ���һ����������󣿣���
    '�����Ч����������ص�ǰȡ��������µ����ж������б����ȥ��
    '����һЩ���水ť���Ƶȡ�
    Debug.Print "txtOutWarehouseReceiptNo_LostFocus"
    Call cmdSearch_Click
End Sub

