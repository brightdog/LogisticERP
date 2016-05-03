VERSION 5.00
Begin VB.Form frmSuppliers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suppliers"
   ClientHeight    =   5085
   ClientLeft      =   2175
   ClientTop       =   2595
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSuppliers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   339
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   3
      Left            =   5130
      TabIndex        =   4
      Top             =   1485
      Width           =   2340
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3810
      TabIndex        =   13
      Top             =   4635
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2505
      TabIndex        =   12
      Top             =   4635
      Width           =   1200
   End
   Begin VB.TextBox txtField 
      Height          =   720
      Index           =   4
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1905
      Width           =   6180
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   9
      Left            =   5040
      TabIndex        =   11
      Top             =   3765
      Width           =   2400
   End
   Begin VB.ComboBox cboCountry 
      Height          =   315
      Left            =   5040
      TabIndex        =   9
      Top             =   3315
      Width           =   2400
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   6
      Left            =   5040
      TabIndex        =   7
      Top             =   2850
      Width           =   2400
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   8
      Left            =   1305
      TabIndex        =   10
      Top             =   3720
      Width           =   2400
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   7
      Left            =   1305
      TabIndex        =   8
      Top             =   3255
      Width           =   2400
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   5
      Left            =   1305
      TabIndex        =   6
      Top             =   2820
      Width           =   2400
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   2
      Left            =   1305
      TabIndex        =   3
      Top             =   1485
      Width           =   3150
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   1
      Left            =   1305
      TabIndex        =   1
      Top             =   1080
      Width           =   6180
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   660
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frmSuppliers.frx":030A
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000003&
      Caption         =   "Suppliers "
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
      Height          =   540
      Left            =   -75
      TabIndex        =   24
      Top             =   0
      Width           =   7635
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   195
      Index           =   2
      Left            =   4710
      TabIndex        =   23
      Top             =   1560
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   -2
      X2              =   503
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   -2
      X2              =   503
      Y1              =   299
      Y2              =   299
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   195
      Index           =   4
      Left            =   600
      TabIndex        =   22
      Top             =   1965
      Width           =   645
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Fax:"
      Height          =   195
      Index           =   10
      Left            =   4320
      TabIndex        =   21
      Top             =   3810
      Width           =   330
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Phone:"
      Height          =   195
      Index           =   9
      Left            =   735
      TabIndex        =   20
      Top             =   3795
      Width           =   510
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Country:"
      Height          =   195
      Index           =   8
      Left            =   4005
      TabIndex        =   19
      Top             =   3345
      Width           =   645
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Postal Code:"
      Height          =   195
      Index           =   7
      Left            =   330
      TabIndex        =   18
      Top             =   3330
      Width           =   915
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Region:"
      Height          =   195
      Index           =   6
      Left            =   4095
      TabIndex        =   17
      Top             =   2895
      Width           =   555
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "City:"
      Height          =   195
      Index           =   5
      Left            =   900
      TabIndex        =   16
      Top             =   2880
      Width           =   345
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Contact Name:"
      Height          =   195
      Index           =   3
      Left            =   165
      TabIndex        =   15
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Company Name:"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   14
      Top             =   1125
      Width           =   1185
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Supplier ID:"
      Height          =   195
      Index           =   0
      Left            =   405
      TabIndex        =   2
      Top             =   735
      Width           =   840
   End
End
Attribute VB_Name = "frmSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_conn As Connection
Dim mrstSuppliers As Recordset
Dim mvarBookmark As Variant
Dim mbIsNew As Boolean
Public Key As String

Private Const fldSupplierID = 0
Private Const fldCompanyName = 1
Private Const fldContactName = 2
Private Const fldContactTitle = 3
Private Const fldAddress = 4
Private Const fldCity = 5
Private Const fldRegion = 6
Private Const fldPostalCode = 7
Private Const fldPhone = 8
Private Const fldFax = 9
Private Const fldCountry = 10
Dim m_DataChanged(10) As Boolean
Public Sub EditRecord(cn As Connection, rs As Recordset)
Dim i As Long

On Error Resume Next
    Set m_conn = cn
    Set mrstSuppliers = rs.Clone
    mvarBookmark = rs.Bookmark
    mrstSuppliers.Bookmark = mvarBookmark
    FillCountryCombo
    txtField(fldSupplierID).Text = mrstSuppliers![SupplierID]
    txtField(fldCompanyName).Text = mrstSuppliers![CompanyName]
    txtField(fldContactName).Text = mrstSuppliers![ContactName]
    txtField(fldContactTitle).Text = mrstSuppliers![ContactTitle]
    txtField(fldAddress).Text = mrstSuppliers![Address]
    txtField(fldCity).Text = mrstSuppliers![City]
    txtField(fldRegion).Text = mrstSuppliers![Region]
    txtField(fldPostalCode).Text = mrstSuppliers![PostalCode]
    txtField(fldPhone).Text = mrstSuppliers![Phone]
    txtField(fldFax).Text = mrstSuppliers![Fax]
    cboCountry.Text = mrstSuppliers![Country]
    Caption = "Suppliers - " & mrstSuppliers![CompanyName]
    For i = 0 To UBound(m_DataChanged)
        m_DataChanged(i) = False
    Next
    Me.Show
End Sub

Public Sub NewRecord(cn As Connection, rs As Recordset)

On Error Resume Next
    Set m_conn = cn
    Set mrstSuppliers = rs.Clone
    mbIsNew = True
    mvarBookmark = Null
    FillCountryCombo
    Caption = "Suppliers - " & mrstSuppliers![CompanyName]
    Me.Show
End Sub


Private Sub cboCountry_Click()

    m_DataChanged(fldCountry) = True
    
End Sub

Private Sub cboCountry_Change()

    m_DataChanged(fldCountry) = True
    
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
            mrstSuppliers.AddNew
        End If
        If m_DataChanged(fldSupplierID) Then mrstSuppliers![SupplierID] = txtField(fldSupplierID)
        If m_DataChanged(fldCompanyName) Then mrstSuppliers![CompanyName] = txtField(fldCompanyName)
        If m_DataChanged(fldContactName) Then mrstSuppliers![ContactName] = txtField(fldContactName)
        If m_DataChanged(fldContactTitle) Then mrstSuppliers![ContactTitle] = txtField(fldContactTitle)
        If m_DataChanged(fldAddress) Then mrstSuppliers![Address] = txtField(fldAddress)
        If m_DataChanged(fldCity) Then mrstSuppliers![City] = txtField(fldCity)
        If m_DataChanged(fldRegion) Then mrstSuppliers![Region] = txtField(fldRegion)
        If m_DataChanged(fldPostalCode) Then mrstSuppliers![PostalCode] = txtField(fldPostalCode)
        If m_DataChanged(fldPhone) Then mrstSuppliers![Phone] = txtField(fldPhone)
        If m_DataChanged(fldFax) Then mrstSuppliers![Fax] = txtField(fldFax)
        If m_DataChanged(fldCountry) Then mrstSuppliers![Country] = cboCountry.Text
        mrstSuppliers.Update
        Hide
        frmMain.OnRecordUpdate CatalogSuppliers, mvarBookmark
    End If
    Unload Me

    Exit Sub
    
EH_cmdOK:
    MsgBox Err.Description

End Sub


Private Sub Form_Unload(Cancel As Integer)

    frmMain.UnloadForm Key
End Sub


Private Sub txtField_Change(Index As Integer)

    m_DataChanged(Index) = True
    
End Sub



Private Sub FillCountryCombo()
Dim rstCountry As Recordset
    Set rstCountry = New Recordset
    rstCountry.Open "SELECT DISTINCT Suppliers.Country FROM Suppliers ORDER BY Suppliers.Country", m_conn, adOpenForwardOnly, adLockReadOnly
    cboCountry.Clear
    cboCountry.AddItem ""
    Do Until rstCountry.EOF
        cboCountry.AddItem rstCountry![Country]
        rstCountry.MoveNext
    Loop
    
End Sub
