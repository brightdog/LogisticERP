VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.0#0"; "GridEX20.ocx"
Begin VB.Form frmOrders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orders"
   ClientHeight    =   5460
   ClientLeft      =   1470
   ClientTop       =   3045
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOrders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8055
   Begin VB.PictureBox Picture2 
      Height          =   1425
      Left            =   4020
      ScaleHeight     =   1365
      ScaleWidth      =   3885
      TabIndex        =   32
      Top             =   435
      Width           =   3945
      Begin VB.TextBox txtShip 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   840
         TabIndex        =   38
         Top             =   30
         Width           =   3000
      End
      Begin VB.TextBox txtShip 
         BorderStyle     =   0  'None
         Height          =   405
         Index           =   1
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   315
         Width           =   3000
      End
      Begin VB.TextBox txtShip 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   2865
         TabIndex        =   36
         Top             =   780
         Width           =   975
      End
      Begin VB.TextBox txtShip 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   5
         Left            =   1710
         TabIndex        =   35
         Top             =   1080
         Width           =   2130
      End
      Begin VB.TextBox txtShip 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   840
         TabIndex        =   34
         Top             =   780
         Width           =   975
      End
      Begin VB.TextBox txtShip 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   1852
         TabIndex        =   33
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ship To:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   75
         TabIndex        =   39
         Top             =   60
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1785
      Left            =   90
      ScaleHeight     =   1725
      ScaleWidth      =   3825
      TabIndex        =   22
      Top             =   75
      Width           =   3885
      Begin VB.TextBox txtCustInfo 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1140
         Width           =   975
      End
      Begin VB.TextBox txtCustInfo 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1140
         Width           =   975
      End
      Begin VB.TextBox txtCustInfo 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   5
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1425
         Width           =   2130
      End
      Begin VB.TextBox txtCustInfo 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1140
         Width           =   975
      End
      Begin VB.TextBox txtCustInfo 
         BorderStyle     =   0  'None
         Height          =   405
         Index           =   1
         Left            =   780
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   690
         Width           =   3000
      End
      Begin VB.TextBox txtCustInfo 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   405
         Width           =   3000
      End
      Begin VB.CommandButton cmdCustomerID 
         Caption         =   "..."
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   75
         Width           =   345
      End
      Begin VB.TextBox txtCustomerID 
         Height          =   315
         Left            =   780
         TabIndex        =   24
         Top             =   45
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bill To:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   75
         TabIndex        =   31
         Top             =   75
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   3015
      TabIndex        =   7
      Top             =   5040
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   4335
      TabIndex        =   8
      Top             =   5040
      Width           =   1200
   End
   Begin VB.ComboBox cboShippers 
      Height          =   315
      Left            =   5310
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.ComboBox cboEmployee 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   2805
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Index           =   8
      Left            =   6660
      TabIndex        =   5
      Top             =   2250
      Width           =   1305
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Index           =   7
      Left            =   3960
      TabIndex        =   4
      Top             =   2250
      Width           =   1200
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Index           =   6
      Left            =   1215
      TabIndex        =   3
      Top             =   2250
      Width           =   1200
   End
   Begin VB.TextBox txtFreight 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6615
      TabIndex        =   10
      Top             =   4777
      Visible         =   0   'False
      Width           =   1290
   End
   Begin GridEX20.GridEX gexDetails 
      Height          =   1695
      Left            =   90
      TabIndex        =   6
      Top             =   2655
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   2990
      Version         =   "2.0"
      ColumnAutoResize=   -1  'True
      HeaderStyle     =   3
      MethodHoldFields=   -1  'True
      SelectionStyle  =   1
      Options         =   8
      RecordsetType   =   1
      AutomaticArrange=   0   'False
      AllowDelete     =   -1  'True
      BorderStyle     =   2
      GroupByBoxVisible=   0   'False
      NewRowPos       =   1
      RowHeaders      =   -1  'True
      DataMode        =   1
      HeaderFontBold  =   -1  'True
      HeaderFontWeight=   700
      AllowAddNew     =   -1  'True
      ColumnHeaderHeight=   285
      IntProp8        =   0   'False
      ColumnsCount    =   7
      Column(1)       =   "frmOrders.frx":014A
      Column(2)       =   "frmOrders.frx":02AA
      Column(3)       =   "frmOrders.frx":0426
      Column(4)       =   "frmOrders.frx":05CE
      Column(5)       =   "frmOrders.frx":0746
      Column(6)       =   "frmOrders.frx":08BE
      Column(7)       =   "frmOrders.frx":0A12
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmOrders.frx":0B9E
      FormatStyle(2)  =   "frmOrders.frx":0CC6
      FormatStyle(3)  =   "frmOrders.frx":0D76
      FormatStyle(4)  =   "frmOrders.frx":0E2A
      FormatStyle(5)  =   "frmOrders.frx":0F02
      ImageCount      =   0
      PrinterProperties=   "frmOrders.frx":0FBA
   End
   Begin VB.Label lblFreight 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   6615
      TabIndex        =   19
      Top             =   4777
      Width           =   1290
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   6615
      TabIndex        =   20
      Top             =   5077
      Width           =   1290
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   6615
      TabIndex        =   21
      Top             =   4477
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   5685
      TabIndex        =   13
      Top             =   5100
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Freight:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5685
      TabIndex        =   12
      Top             =   4800
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SubTotal:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5685
      TabIndex        =   11
      Top             =   4500
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ship Via:"
      Height          =   195
      Index           =   4
      Left            =   4530
      TabIndex        =   18
      Top             =   1965
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Salesperson:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1995
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Shipped Date:"
      Height          =   195
      Index           =   2
      Left            =   5385
      TabIndex        =   16
      Top             =   2310
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Required Date:"
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   15
      Top             =   2310
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Order Date:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   2310
      Width           =   870
   End
   Begin VB.Label lblOrderNo 
      BackStyle       =   0  'Transparent
      Caption         =   " N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7335
      TabIndex        =   9
      Top             =   75
      Width           =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDER #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   6405
      TabIndex        =   0
      Top             =   105
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   3
      Left            =   6060
      Picture         =   "frmOrders.frx":118A
      Stretch         =   -1  'True
      Top             =   75
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   5970
      Top             =   15
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   5595
      TabIndex        =   40
      Top             =   4410
      Width           =   2400
   End
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_conn As Connection
Dim mrstOrders As Recordset
Dim mrstCustomers As Recordset
Dim mrstProducts As Recordset

Dim mbIsNew As Boolean
Dim mvarBookmark As Variant

Const fldShipTo = 0
Const fldShipAddress = 1
Const fldShipCity = 2
Const fldShipRegion = 3
Const fldShipPostalCode = 4
Const fldShipCountry = 5
Const fldOrderDate = 6
Const fldRequiredDate = 7
Const fldShippedDate = 8
Const fldEmployeeID = 9
Const fldShipVia = 10
Const fldFreight = 11
Const fldCustomerID = 12

Const fldCustName = 0
Const fldCustAddress = 1
Const fldCustCity = 2
Const fldCustRegion = 3
Const fldCustPostalCode = 4
Const fldCustCountry = 5

Dim m_DataChanged(0 To 12) As Boolean

Public Key As String


Private Sub CalculateTotals(Freight As Currency)
Dim rst As Recordset
Dim amount As Currency
    Set rst = gexDetails.adoRecordset
    On Error Resume Next
    rst.MoveFirst
    Do Until rst.EOF
        amount = amount + rst![Price]
        rst.MoveNext
    Loop
    lblSubTotal = Format(amount, "Currency")
    lblFreight = Format(Freight, "Currency")
    lblTotal = Format(amount + Freight, "Currency")
    
End Sub


Private Sub cboEmployee_Click()
    m_DataChanged(fldEmployeeID) = True

End Sub

Private Sub cboEmployee_Change()

    m_DataChanged(fldEmployeeID) = True
End Sub

Private Sub cboShippers_Click()
    m_DataChanged(fldShipVia) = True

End Sub

Private Sub cboShippers_Change()
    m_DataChanged(fldShipVia) = True
    
End Sub

Private Sub cmdCancel_Click()

    If mbIsNew And Not IsNull(mvarBookmark) Then
        mrstOrders.Delete
    End If
    Unload Me
End Sub

Private Sub cmdCustomerID_Click()
Dim varCustID As Variant

    varCustID = frmList.ChooseCustomer(m_conn.ConnectionString, txtCustomerID)
    If Not IsNull(varCustID) Then
        txtCustomerID = varCustID
        txtCustomerID.SelStart = 0
        txtCustomerID.SelLength = Len(txtCustomerID)
    End If
    txtCustomerID.SetFocus
    
End Sub


Private Sub cmdOK_Click()
On Error GoTo EH_cmdOK

    If ActiveControl Is txtFreight Then
        txtFreight_LostFocus
    ElseIf ActiveControl Is gexDetails Then
        gexDetails.Update
    End If
    If Not SaveOrder Then Exit Sub
    If mbIsNew Then
        frmMain.OnRecordUpdate CatalogOrders, Null
    Else
        frmMain.OnRecordUpdate CatalogOrders, mvarBookmark
    End If
    
    Unload Me
    Exit Sub
    
EH_cmdOK:
    MsgBox Err.Description

End Sub







Private Sub Form_Unload(Cancel As Integer)
    frmMain.UnloadForm Key
    
End Sub


Private Sub gexDetails_AfterColUpdate(ByVal ColIndex As Integer)
On Error Resume Next
Dim curUnitPrice As Currency
Dim sngQuantity As Single
Dim sngDiscount As Single
Dim colQuantity As JSColumn

    Set colQuantity = gexDetails.Columns("Quantity")
    If Not colQuantity.DataChanged Then
        gexDetails.Value(colQuantity.Index) = 1
    End If
    Select Case ColIndex
        Case 2
            mrstProducts.MoveFirst
            mrstProducts.Find "ProductID=" & gexDetails.Value(2)
            If mrstProducts.EOF Then
                gexDetails.Value(6) = ""
            Else
                gexDetails.Value(6) = mrstProducts![ProductName]
                gexDetails.Value(3) = mrstProducts![UnitPrice]
                gexDetails_AfterColUpdate 3
            End If
        Case 3, 4, 5
            curUnitPrice = CCur(gexDetails.Value(3))
            sngQuantity = CSng(gexDetails.Value(4))
            sngDiscount = CSng(gexDetails.Value(5))
            gexDetails.Value(7) = curUnitPrice * sngQuantity * (1 - sngDiscount)
    End Select
    
End Sub

Private Sub gexDetails_AfterDelete()

    gexDetails_AfterUpdate
End Sub

Private Sub gexDetails_AfterUpdate()

    If lblFreight.Caption = "" Then
        CalculateTotals 0
    Else
        CalculateTotals CCur(lblFreight.Caption)
    End If
End Sub


Private Sub gexDetails_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)

    If IsNull(mvarBookmark) Then
        If Len(txtCustomerID.Text) = 0 Then
            MsgBox "Select a customer to bill to before entering order details info.", vbInformation
            gexDetails.DataChanged = False
            txtCustomerID.SetFocus
            Cancel = True
            Exit Sub
        Else
            If Not SaveOrder Then
                Cancel = True
            End If
        End If
    End If
    If gexDetails.Value(1) = "" Then
        gexDetails.Value(1) = mrstOrders![OrderID]
    End If
    gexDetails.Columns(6).DataChanged = False
    gexDetails.Columns(7).DataChanged = False
    
End Sub

Private Sub lblFreight_Click()
    txtFreight.Text = lblFreight.Caption
    txtFreight.SelStart = 0
    txtFreight.SelLength = Len(txtFreight)
    txtFreight.Visible = True
    txtFreight.SetFocus
End Sub

Private Sub txtCustomerID_Change()
    SearchCustomer
    m_DataChanged(fldCustomerID) = True
    
End Sub



Private Sub SearchCustomer()
On Error Resume Next
Dim i As Integer
    mrstCustomers.Find "CustomerID ='" & txtCustomerID & "'"
    If mrstCustomers.EOF Then
        For i = fldCustName To fldCustCountry
            txtCustInfo(i).Text = ""
        Next
    Else
        txtCustInfo(fldCustName) = mrstCustomers![CompanyName]
        txtCustInfo(fldCustAddress) = mrstCustomers![Address]
        txtCustInfo(fldCustCity) = mrstCustomers![City]
        txtCustInfo(fldCustRegion) = mrstCustomers![Region]
        txtCustInfo(fldCustPostalCode) = mrstCustomers![PostalCode]
        txtCustInfo(fldCustCountry) = mrstCustomers![Country]
        If Not m_DataChanged(fldShipTo) Then txtShip(fldShipTo).Text = mrstCustomers![ContactName]
        If Not m_DataChanged(fldShipAddress) Then txtShip(fldShipAddress).Text = mrstCustomers![Address]
        If Not m_DataChanged(fldShipCity) Then txtShip(fldShipCity).Text = mrstCustomers![City]
        If Not m_DataChanged(fldShipRegion) Then txtShip(fldShipRegion).Text = mrstCustomers![Region]
        If Not m_DataChanged(fldShipPostalCode) Then txtShip(fldShipPostalCode).Text = mrstCustomers![PostalCode]
        If Not m_DataChanged(fldShipCountry) Then txtShip(fldShipCountry).Text = mrstCustomers![Country]
    End If
End Sub
Public Sub EditRecord(cn As Connection, rs As Recordset)
Dim i As Long
Dim lngID As Long

On Error Resume Next
    Set m_conn = cn
    Set mrstCustomers = New Recordset
    mrstCustomers.Open "SELECT * FROM Customers", m_conn, adOpenStatic, adLockReadOnly
    Set mrstProducts = New Recordset
    mrstProducts.Open "SELECT * FROM Products", m_conn, adOpenStatic, adLockReadOnly
    Set mrstOrders = rs.Clone
    mvarBookmark = rs.Bookmark
    mrstOrders.Bookmark = mvarBookmark
    FillEmployeeList
    FillShippersList
    txtCustomerID.Text = mrstOrders![CustomerID]
    txtShip(fldShipTo).Text = mrstOrders![ShipName]
    txtShip(fldShipAddress).Text = mrstOrders![ShipAddress]
    txtShip(fldShipCity).Text = mrstOrders![ShipCity]
    txtShip(fldShipRegion).Text = mrstOrders![ShipRegion]
    txtShip(fldShipPostalCode).Text = mrstOrders![ShipPostalCode]
    txtShip(fldShipCountry).Text = mrstOrders![ShipCountry]
    txtDate(fldOrderDate) = Format(mrstOrders![OrderDate], "Medium Date")
    txtDate(fldRequiredDate) = Format(mrstOrders![RequiredDate], "Medium Date")
    txtDate(fldShippedDate) = Format(mrstOrders![ShippedDate], "Medium Date")
    If Not IsNull(mrstOrders![EmployeeID]) Then
        lngID = mrstOrders![EmployeeID]
    End If
    For i = 0 To cboEmployee.ListCount - 1
        If cboEmployee.ItemData(i) = lngID Then
            cboEmployee.ListIndex = i
            Exit For
        End If
    Next
    If Not IsNull(mrstOrders![ShipVia]) Then
        lngID = mrstOrders![ShipVia]
    Else
        lngID = 0
    End If
    For i = 0 To cboShippers.ListCount - 1
        If cboShippers.ItemData(i) = lngID Then
            cboShippers.ListIndex = i
            Exit For
        End If
    Next
    lblOrderNo.Caption = mrstOrders![OrderID]
    Caption = "Orders - Order # " & mrstOrders![OrderID]
    For i = 0 To UBound(m_DataChanged)
        m_DataChanged(i) = False
    Next
    gexDetails.DatabaseName = m_conn.ConnectionString
    gexDetails.RecordSource = "SELECT [Order Details].*, Products.ProductName, ([Order Details]![UnitPrice]*[Order Details]![Quantity])*(1-[Order Details]![Discount]) AS Price FROM Products INNER JOIN [Order Details] ON Products.ProductID = [Order Details].ProductID WHERE [Order Details].OrderID=" & mrstOrders![OrderID]
    gexDetails.HoldFields
    gexDetails.Rebind
    If IsNull(mrstOrders![Freight]) Then
        Call CalculateTotals(0)
    Else
        Call CalculateTotals(mrstOrders![Freight])
    End If
    Show
End Sub


Public Sub NewRecord(cn As Connection, rs As Recordset)
Dim i As Long
Dim lngID As Long

On Error Resume Next
    Set m_conn = cn
    Set mrstCustomers = New Recordset
    mrstCustomers.Open "SELECT * FROM Customers", m_conn, adOpenStatic, adLockReadOnly
    Set mrstProducts = New Recordset
    mrstProducts.Open "SELECT * FROM Products", m_conn, adOpenStatic, adLockReadOnly
    Set mrstOrders = rs.Clone
    mbIsNew = True
    mvarBookmark = Null
    FillEmployeeList
    FillShippersList
    txtDate(fldOrderDate) = Format(Date, "Medium Date")
    txtDate(fldShippedDate) = Format(Date, "Medium Date")
    txtDate(fldRequiredDate) = Format(Date, "Medium Date")
    Caption = "Orders - New Order"
    For i = 0 To UBound(m_DataChanged)
        m_DataChanged(i) = False
    Next
    gexDetails.DatabaseName = m_conn.ConnectionString
    gexDetails.RecordSource = "SELECT [Order Details].*, Products.ProductName, ([Order Details]![UnitPrice]*[Order Details]![Quantity])*(1-[Order Details]![Discount]) AS Price FROM Products INNER JOIN [Order Details] ON Products.ProductID = [Order Details].ProductID WHERE [Order Details].OrderID=0"
    gexDetails.HoldFields
    gexDetails.Rebind
    Call CalculateTotals(0)
    Me.Show
End Sub








Private Sub FillEmployeeList()
Dim rsTemp As Recordset

    Set rsTemp = New Recordset
    
    rsTemp.Open "SELECT Employees.EmployeeID, Employees.FirstName & ' ' & Employees.LastName AS Name From Employees", m_conn, adOpenForwardOnly, adLockReadOnly
    cboEmployee.Clear
    cboEmployee.AddItem ""
    Do Until rsTemp.EOF
        cboEmployee.AddItem rsTemp![Name]
        cboEmployee.ItemData(cboEmployee.NewIndex) = rsTemp![EmployeeID]
        rsTemp.MoveNext
    Loop
    

End Sub
Private Sub FillShippersList()
Dim rsTemp As Recordset

    Set rsTemp = New Recordset
    rsTemp.Open "SELECT * From Shippers", m_conn, adOpenForwardOnly, adLockReadOnly
    cboShippers.Clear
    cboShippers.AddItem ""
    Do Until rsTemp.EOF
        cboShippers.AddItem rsTemp![CompanyName]
        cboShippers.ItemData(cboShippers.NewIndex) = rsTemp![ShipperID]
        rsTemp.MoveNext
    Loop
    

End Sub

Private Sub txtDate_Change(Index As Integer)
    m_DataChanged(Index) = True
End Sub

Private Sub txtFreight_Change()

    m_DataChanged(fldFreight) = True
    
End Sub

Private Sub txtFreight_LostFocus()
Dim cTemp As Currency
On Error Resume Next
    cTemp = CCur(txtFreight)
    If Err = 0 Then
        CalculateTotals (cTemp)
    End If
    txtFreight.Visible = False
    
End Sub


Private Sub txtShip_Change(Index As Integer)

    m_DataChanged(Index) = True
    
End Sub



Private Function SaveOrder() As Boolean
On Error GoTo EH_SaveOrder
Dim curTemp As Currency
    If IsDirty Then
        If IsNull(mvarBookmark) Then
            mrstOrders.AddNew
        Else
            mrstOrders.Bookmark = mvarBookmark
        End If
        If m_DataChanged(fldCustomerID) Then mrstOrders![CustomerID] = txtCustomerID
        If m_DataChanged(fldShipTo) Then mrstOrders![ShipName] = txtShip(fldShipTo)
        If m_DataChanged(fldShipAddress) Then mrstOrders![ShipAddress] = txtShip(fldShipAddress)
        If m_DataChanged(fldShipRegion) Then mrstOrders![ShipRegion] = txtShip(fldShipRegion)
        If m_DataChanged(fldShipCity) Then mrstOrders![ShipCity] = txtShip(fldShipCity)
        If m_DataChanged(fldShipPostalCode) Then mrstOrders![ShipPostalCode] = txtShip(fldShipPostalCode)
        If m_DataChanged(fldShipCountry) Then mrstOrders![ShipCountry] = txtShip(fldShipCountry)
        If m_DataChanged(fldOrderDate) Then mrstOrders![OrderDate] = TextToNull(txtDate(fldOrderDate))
        If m_DataChanged(fldRequiredDate) Then mrstOrders![RequiredDate] = TextToNull(txtDate(fldRequiredDate))
        If m_DataChanged(fldShippedDate) Then mrstOrders![ShippedDate] = TextToNull(txtDate(fldShippedDate))
        If m_DataChanged(fldFreight) Then
            On Error Resume Next
            curTemp = CCur(lblFreight)
            On Error GoTo EH_SaveOrder
            mrstOrders![Freight] = curTemp
        End If
        If m_DataChanged(fldEmployeeID) Then
            If cboEmployee.Text = "" Then
                mrstOrders![EmployeeID] = Null
            Else
                mrstOrders![EmployeeID] = cboEmployee.ItemData(cboEmployee.ListIndex)
            End If
        End If
        If m_DataChanged(fldShipVia) Then
            If cboShippers.Text = "" Then
                mrstOrders![ShipVia] = Null
            Else
                mrstOrders![ShipVia] = cboShippers.ItemData(cboShippers.ListIndex)
            End If
        End If
        Dim vTemp As Variant
        mrstOrders.Update
        lblOrderNo.Caption = mrstOrders![OrderID]
        Caption = "Orders - Order # " & mrstOrders![OrderID]
    End If
    SaveOrder = True
    Exit Function
    
EH_SaveOrder:
    MsgBox Err.Description, vbExclamation
End Function

Private Function IsDirty() As Boolean
Dim i As Integer
    For i = 0 To UBound(m_DataChanged)
        If m_DataChanged(i) Then
            IsDirty = True
            Exit Function
        End If
    Next
End Function
