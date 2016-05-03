VERSION 5.00
Begin VB.Form frmCustomers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers"
   ClientHeight    =   4875
   ClientLeft      =   1485
   ClientTop       =   2370
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   3
      Left            =   5100
      TabIndex        =   3
      Top             =   1515
      Width           =   2340
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3840
      TabIndex        =   12
      Top             =   4470
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2535
      TabIndex        =   11
      Top             =   4470
      Width           =   1200
   End
   Begin VB.TextBox txtField 
      Height          =   720
      Index           =   4
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1935
      Width           =   6180
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   9
      Left            =   5100
      TabIndex        =   10
      Top             =   3750
      Width           =   1800
   End
   Begin VB.ComboBox cboCountry 
      Height          =   315
      Left            =   5100
      TabIndex        =   8
      Top             =   3300
      Width           =   1800
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   6
      Left            =   5100
      TabIndex        =   6
      Top             =   2835
      Width           =   1800
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   8
      Left            =   1260
      TabIndex        =   9
      Top             =   3750
      Width           =   2340
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   7
      Left            =   1260
      TabIndex        =   7
      Top             =   3285
      Width           =   2340
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   5
      Left            =   1260
      TabIndex        =   5
      Top             =   2835
      Width           =   2340
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   2
      Left            =   1260
      TabIndex        =   2
      Top             =   1515
      Width           =   3150
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Top             =   1110
      Width           =   6180
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   690
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmCustomers.frx":014A
      Top             =   30
      Width           =   480
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000003&
      Caption         =   "Customers "
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
      Height          =   525
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   7710
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   195
      Index           =   2
      Left            =   4680
      TabIndex        =   23
      Top             =   1590
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   505
      Y1              =   289
      Y2              =   289
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   505
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   195
      Index           =   4
      Left            =   555
      TabIndex        =   22
      Top             =   1995
      Width           =   645
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Fax:"
      Height          =   195
      Index           =   10
      Left            =   4710
      TabIndex        =   21
      Top             =   3825
      Width           =   330
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Phone:"
      Height          =   195
      Index           =   9
      Left            =   705
      TabIndex        =   20
      Top             =   3825
      Width           =   510
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Country:"
      Height          =   195
      Index           =   8
      Left            =   4395
      TabIndex        =   19
      Top             =   3360
      Width           =   645
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Postal Code:"
      Height          =   195
      Index           =   7
      Left            =   300
      TabIndex        =   18
      Top             =   3360
      Width           =   915
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Region:"
      Height          =   195
      Index           =   6
      Left            =   4485
      TabIndex        =   17
      Top             =   2910
      Width           =   555
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "City:"
      Height          =   195
      Index           =   5
      Left            =   870
      TabIndex        =   16
      Top             =   2910
      Width           =   345
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Contact Name:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1590
      Width           =   1080
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Company Name:"
      Height          =   195
      Index           =   1
      Left            =   15
      TabIndex        =   14
      Top             =   1155
      Width           =   1185
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Customer ID:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   765
      Width           =   960
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_conn As Connection
Dim mrstCustomers As Recordset
Dim mvarBookmark As Variant
Dim mbIsNew As Boolean
Public Key As String

Private Const fldCustomerID = 0
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
    Set mrstCustomers = rs.Clone
    mvarBookmark = rs.Bookmark
    mrstCustomers.Bookmark = mvarBookmark
    FillCountryCombo
    txtField(fldCustomerID).Text = mrstCustomers![CustomerID]
    txtField(fldCompanyName).Text = mrstCustomers![CompanyName]
    txtField(fldContactName).Text = mrstCustomers![ContactName]
    txtField(fldContactTitle).Text = mrstCustomers![ContactTitle]
    txtField(fldAddress).Text = mrstCustomers![Address]
    txtField(fldCity).Text = mrstCustomers![City]
    txtField(fldRegion).Text = mrstCustomers![Region]
    txtField(fldPostalCode).Text = mrstCustomers![PostalCode]
    txtField(fldPhone).Text = mrstCustomers![Phone]
    txtField(fldFax).Text = mrstCustomers![Fax]
    cboCountry.Text = mrstCustomers![Country]
    Caption = "Customers - " & mrstCustomers![CompanyName]
    For i = 0 To UBound(m_DataChanged)
        m_DataChanged(i) = False
    Next
    Me.Show
End Sub

Public Sub NewRecord(cn As Connection, rs As Recordset)

On Error Resume Next
    Set m_conn = cn
    Set mrstCustomers = rs.Clone
    mbIsNew = True
    mvarBookmark = Null
    FillCountryCombo
    Caption = "Customers - New Customer"
    Me.Show
End Sub
Private Sub FillCountryCombo()
Dim rstCountry As Recordset
    Set rstCountry = New Recordset
    rstCountry.Open "SELECT DISTINCT Customers.Country FROM Customers ORDER BY Customers.Country", m_conn, adOpenForwardOnly, adLockReadOnly
    cboCountry.Clear
    cboCountry.AddItem ""
    Do Until rstCountry.EOF
        If Not IsNull(rstCountry![Country]) Then cboCountry.AddItem rstCountry![Country]
        rstCountry.MoveNext
    Loop

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
            mrstCustomers.AddNew
        End If
        If m_DataChanged(fldCustomerID) Then mrstCustomers![CustomerID] = txtField(fldCustomerID)
        If m_DataChanged(fldCompanyName) Then mrstCustomers![CompanyName] = txtField(fldCompanyName)
        If m_DataChanged(fldContactName) Then mrstCustomers![ContactName] = txtField(fldContactName)
        If m_DataChanged(fldContactTitle) Then mrstCustomers![ContactTitle] = txtField(fldContactTitle)
        If m_DataChanged(fldAddress) Then mrstCustomers![Address] = txtField(fldAddress)
        If m_DataChanged(fldCity) Then mrstCustomers![City] = txtField(fldCity)
        If m_DataChanged(fldRegion) Then mrstCustomers![Region] = txtField(fldRegion)
        If m_DataChanged(fldPostalCode) Then mrstCustomers![PostalCode] = txtField(fldPostalCode)
        If m_DataChanged(fldPhone) Then mrstCustomers![Phone] = txtField(fldPhone)
        If m_DataChanged(fldFax) Then mrstCustomers![Fax] = txtField(fldFax)
        If m_DataChanged(fldCountry) Then mrstCustomers![Country] = cboCountry.Text
        mrstCustomers.Update
        Hide
        frmMain.OnRecordUpdate CatalogCustomers, mvarBookmark
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


