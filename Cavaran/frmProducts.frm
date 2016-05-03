VERSION 5.00
Begin VB.Form frmProducts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Products"
   ClientHeight    =   5100
   ClientLeft      =   1200
   ClientTop       =   2565
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProducts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   Begin VB.CheckBox chkOnsale 
      Alignment       =   1  'Right Justify
      Caption         =   "On Sale:"
      Height          =   255
      Left            =   3900
      TabIndex        =   11
      Top             =   4050
      Width           =   990
   End
   Begin VB.CheckBox chkDiscontinued 
      Alignment       =   1  'Right Justify
      Caption         =   "Discontinued:"
      Height          =   255
      Left            =   330
      TabIndex        =   10
      Top             =   4020
      Width           =   1365
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   3
      Left            =   1470
      TabIndex        =   5
      Top             =   2415
      Width           =   4830
   End
   Begin VB.CommandButton cmdSuplierList 
      Caption         =   "..."
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1650
      Width           =   300
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3570
      TabIndex        =   13
      Top             =   4635
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2265
      TabIndex        =   12
      Top             =   4635
      Width           =   1200
   End
   Begin VB.ComboBox cboCategory 
      Height          =   315
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2025
      Width           =   2625
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   6
      Left            =   4680
      TabIndex        =   7
      Top             =   2970
      Width           =   1620
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   8
      Left            =   1470
      TabIndex        =   8
      Top             =   3495
      Width           =   1620
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   7
      Left            =   4680
      TabIndex        =   9
      Top             =   3480
      Width           =   1620
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   5
      Left            =   1470
      TabIndex        =   6
      Top             =   2985
      Width           =   1620
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   2
      Left            =   1470
      TabIndex        =   2
      Top             =   1620
      Width           =   1500
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   1
      Left            =   1470
      TabIndex        =   1
      Top             =   1215
      Width           =   4815
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   0
      Left            =   1470
      TabIndex        =   0
      Top             =   795
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmProducts.frx":014A
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000003&
      Caption         =   "Products "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   600
      Left            =   -405
      TabIndex        =   24
      Top             =   0
      Width           =   7350
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Quantity per Unit:"
      Height          =   195
      Index           =   2
      Left            =   45
      TabIndex        =   23
      Top             =   2490
      Width           =   1305
   End
   Begin VB.Label lblSupplierName 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3150
      TabIndex        =   22
      Top             =   1695
      Width           =   45
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   -25
      X2              =   970
      Y1              =   302
      Y2              =   302
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   -25
      X2              =   970
      Y1              =   301
      Y2              =   301
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Reorder Level:"
      Height          =   195
      Index           =   9
      Left            =   285
      TabIndex        =   21
      Top             =   3555
      Width           =   1065
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Category:"
      Height          =   195
      Index           =   8
      Left            =   615
      TabIndex        =   20
      Top             =   2130
      Width           =   735
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Units on Order:"
      Height          =   195
      Index           =   7
      Left            =   3435
      TabIndex        =   19
      Top             =   3555
      Width           =   1110
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Unit in Stock:"
      Height          =   195
      Index           =   6
      Left            =   3600
      TabIndex        =   18
      Top             =   3000
      Width           =   945
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Unit Price:"
      Height          =   195
      Index           =   5
      Left            =   615
      TabIndex        =   17
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Supplier ID:"
      Height          =   195
      Index           =   3
      Left            =   510
      TabIndex        =   16
      Top             =   1665
      Width           =   840
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Product Name:"
      Height          =   195
      Index           =   1
      Left            =   285
      TabIndex        =   15
      Top             =   1275
      Width           =   1065
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Product ID:"
      Height          =   195
      Index           =   0
      Left            =   525
      TabIndex        =   14
      Top             =   870
      Width           =   825
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_conn As Connection
Dim mrstProducts As Recordset
Dim mrstSuppliers As Recordset
Dim mvarBookmark As Variant
Dim mbIsNew As Boolean
Public Key As String

Private Const fldProductID = 0
Private Const fldProductName = 1
Private Const fldSupplierID = 2
Private Const fldQuantityPerUnit = 3
Private Const fldCategory = 4
Private Const fldUnitPrice = 5
Private Const fldUnitsInStock = 6
Private Const fldUnitsOnOrder = 7
Private Const fldReorderLevel = 8
Private Const fldDiscontinued = 9
Private Const fldOnSale = 10
Dim m_DataChanged(10) As Boolean
Public Sub EditRecord(cn As Connection, rs As Recordset)


Dim i As Long

On Error Resume Next
    Set m_conn = cn
    Set mrstSuppliers = New Recordset
    mrstSuppliers.Open "SELECT * FROM Suppliers", m_conn, adOpenStatic, adLockReadOnly
    FillCategoriesCombo
    Set mrstProducts = rs.Clone
    mvarBookmark = rs.Bookmark
    mrstProducts.Bookmark = mvarBookmark
    
    txtField(fldProductID).Text = mrstProducts![ProductID]
    txtField(fldProductName).Text = mrstProducts![ProductName]
    txtField(fldSupplierID).Text = mrstProducts![SupplierID]
    txtField(fldQuantityPerUnit).Text = mrstProducts![QuantityPerUnit]
    txtField(fldUnitPrice).Text = mrstProducts![UnitPrice]
    txtField(fldUnitsInStock).Text = mrstProducts![UnitsInStock]
    txtField(fldUnitsOnOrder).Text = mrstProducts![UnitsOnOrder]
    txtField(fldReorderLevel).Text = mrstProducts![ReorderLevel]
    If IsNull(mrstProducts![CategoryID]) Then
        cboCategory.ListIndex = 0
    Else
        For i = 1 To cboCategory.ListCount - 1
            If cboCategory.ItemData(i) = mrstProducts![CategoryID] Then
                cboCategory.ListIndex = i
                Exit For
            End If
        Next
    End If
    If Not IsNull(mrstProducts![Discontinued]) Then
        If mrstProducts![Discontinued] Then
            chkDiscontinued.Value = vbChecked
        Else
            chkDiscontinued.Value = vbUnchecked
        End If
    Else
        chkDiscontinued.Value = vbUnchecked
    End If
    If Not IsNull(mrstProducts![OnSale]) Then
        If mrstProducts![OnSale] Then
            chkOnsale.Value = vbChecked
        Else
            chkOnsale.Value = vbUnchecked
        End If
    Else
        chkOnsale.Value = vbUnchecked
    End If

    Caption = "Products - " & mrstProducts![ProductName]
    For i = 0 To UBound(m_DataChanged)
        m_DataChanged(i) = False
    Next
    Me.Show
End Sub


Public Sub NewRecord(cn As Connection, rs As Recordset)
On Error Resume Next
    
    Set m_conn = cn
    Set mrstSuppliers = New Recordset
    mrstSuppliers.Open "SELECT * FROM Suppliers", m_conn, adOpenStatic, adLockOptimistic
    Set mrstProducts = rs.Clone
    mbIsNew = True
    mvarBookmark = Null
    FillCategoriesCombo
    Caption = "Products - New Product"
    Me.Show
End Sub




Private Sub cboCategory_Click()
    m_DataChanged(fldCategory) = True
    
End Sub

Private Sub chkOnsale_Click()

    m_DataChanged(fldOnSale) = True
End Sub

Private Sub cmdCancel_Click()

    Unload Me
End Sub


Private Sub cmdOK_Click()
On Error GoTo EH_cmdOK
Dim bUpdate As Boolean
Dim i As Integer

    For i = 0 To UBound(m_DataChanged)
        If m_DataChanged(i) Then bUpdate = True
    Next
    If bUpdate Then
        If mbIsNew Then
            mrstProducts.AddNew
        End If
        If m_DataChanged(fldProductID) Then mrstProducts![ProductID] = txtField(fldProductID)
        If m_DataChanged(fldProductName) Then mrstProducts![ProductName] = txtField(fldProductName)
        If m_DataChanged(fldSupplierID) Then mrstProducts![SupplierID] = txtField(fldSupplierID)
        If m_DataChanged(fldQuantityPerUnit) Then mrstProducts![QuantityPerUnit] = txtField(fldQuantityPerUnit)
        If m_DataChanged(fldUnitPrice) Then mrstProducts![UnitPrice] = txtField(fldUnitPrice)
        If m_DataChanged(fldUnitsInStock) Then mrstProducts![UnitsInStock] = txtField(fldUnitsInStock)
        If m_DataChanged(fldUnitsOnOrder) Then mrstProducts![UnitsOnOrder] = txtField(fldUnitsOnOrder)
        If m_DataChanged(fldReorderLevel) Then mrstProducts![ReorderLevel] = txtField(fldReorderLevel)
        If m_DataChanged(fldCategory) Then
            If cboCategory.ListIndex <= 0 Then
                mrstProducts![CategoryID] = Null
            Else
                mrstProducts![CategoryID] = cboCategory.ItemData(cboCategory.ListIndex)
            End If
        End If
        If m_DataChanged(fldDiscontinued) Then mrstProducts![Discontinued] = (chkDiscontinued.Value = vbChecked)
        If m_DataChanged(fldOnSale) Then mrstProducts![OnSale] = (chkOnsale.Value = vbChecked)
        mrstProducts.Update
        Hide
        frmMain.OnRecordUpdate CatalogProducts, mvarBookmark
    End If
    Unload Me

    Exit Sub
    
EH_cmdOK:
    MsgBox Err.Description

End Sub


Private Sub cmdSuplierList_Click()
Dim varSup As Variant

    varSup = frmList.ChooseSupplier(m_conn.ConnectionString, txtField(fldSupplierID))
    If Not IsNull(varSup) Then
        txtField(fldSupplierID) = varSup
    End If
    txtField(fldSupplierID).SetFocus
    
End Sub

Private Sub chkDiscontinued_Click()
    m_DataChanged(fldDiscontinued) = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.UnloadForm Key
End Sub


Private Sub txtField_Change(Index As Integer)

    m_DataChanged(Index) = True
    If Index = fldSupplierID Then
        SearchSupplierName
    End If
End Sub



Private Sub SearchSupplierName()
Dim strSupplierID As String

    strSupplierID = txtField(fldSupplierID)
    If strSupplierID = "" Then
        lblSupplierName = ""
    Else
        mrstSuppliers.MoveFirst
        mrstSuppliers.Find "SupplierID=" & strSupplierID
        If mrstSuppliers.EOF Then
            lblSupplierName = ""
        Else
            lblSupplierName = mrstSuppliers![CompanyName]
        End If
    End If
    
End Sub
Private Sub FillCategoriesCombo()
Dim rstCategory As Recordset
    Set rstCategory = New Recordset
    rstCategory.Open "SELECT Categories.CategoryID, Categories.CategoryName FROM Categories ORDER BY Categories.CategoryName", m_conn, adOpenStatic, adLockReadOnly
    cboCategory.Clear
    cboCategory.AddItem ""
    Do Until rstCategory.EOF
        cboCategory.AddItem rstCategory![CategoryName]
        cboCategory.ItemData(cboCategory.NewIndex) = rstCategory![CategoryID]
        rstCategory.MoveNext
    Loop
End Sub
