VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmployees 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employees"
   ClientHeight    =   5535
   ClientLeft      =   3045
   ClientTop       =   1995
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmployees.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   483
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3960
      TabIndex        =   1
      Top             =   5085
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2655
      TabIndex        =   0
      Top             =   5085
      Width           =   1200
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   3705
      Index           =   2
      Left            =   195
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   458
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   6870
      Begin VB.TextBox txtField 
         Height          =   660
         Index           =   12
         Left            =   1110
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   2820
         Width           =   4740
      End
      Begin VB.TextBox txtField 
         Height          =   345
         Index           =   9
         Left            =   1110
         TabIndex        =   20
         Top             =   2382
         Width           =   1200
      End
      Begin VB.TextBox txtField 
         Height          =   345
         Index           =   8
         Left            =   1110
         TabIndex        =   17
         Top             =   1944
         Width           =   2340
      End
      Begin VB.ComboBox cboTitleOfCourtesy 
         Height          =   315
         ItemData        =   "frmEmployees.frx":014A
         Left            =   5025
         List            =   "frmEmployees.frx":015A
         TabIndex        =   16
         Top             =   1935
         Width           =   1800
      End
      Begin VB.ComboBox cboCountry 
         Height          =   315
         Left            =   5025
         TabIndex        =   13
         Top             =   1515
         Width           =   1800
      End
      Begin VB.TextBox txtField 
         Height          =   345
         Index           =   7
         Left            =   1110
         TabIndex        =   12
         Top             =   1506
         Width           =   2340
      End
      Begin VB.TextBox txtField 
         Height          =   345
         Index           =   6
         Left            =   5025
         TabIndex        =   9
         Top             =   1050
         Width           =   1800
      End
      Begin VB.TextBox txtField 
         Height          =   345
         Index           =   5
         Left            =   1110
         TabIndex        =   8
         Top             =   1068
         Width           =   2340
      End
      Begin VB.TextBox txtField 
         Height          =   720
         Index           =   4
         Left            =   1110
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   255
         Width           =   4740
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Notes:"
         Height          =   195
         Index           =   15
         Left            =   510
         TabIndex        =   23
         Top             =   2820
         Width           =   480
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "Birth Date:"
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   21
         Top             =   2430
         Width           =   780
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Home Phone:"
         Height          =   195
         Index           =   9
         Left            =   30
         TabIndex        =   19
         Top             =   2010
         Width           =   960
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "Title of Courtesy:"
         Height          =   195
         Index           =   14
         Left            =   3690
         TabIndex        =   18
         Top             =   2025
         Width           =   1260
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "Country:"
         Height          =   195
         Index           =   8
         Left            =   4305
         TabIndex        =   15
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Postal Code:"
         Height          =   195
         Index           =   7
         Left            =   75
         TabIndex        =   14
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "Region:"
         Height          =   195
         Index           =   6
         Left            =   4395
         TabIndex        =   11
         Top             =   1125
         Width           =   555
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "City:"
         Height          =   195
         Index           =   5
         Left            =   645
         TabIndex        =   10
         Top             =   1125
         Width           =   345
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         Height          =   195
         Index           =   4
         Left            =   345
         TabIndex        =   7
         Top             =   315
         Width           =   645
      End
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   3705
      Index           =   1
      Left            =   165
      ScaleHeight     =   3705
      ScaleWidth      =   6885
      TabIndex        =   4
      Top             =   990
      Width           =   6885
      Begin VB.TextBox txtField 
         Height          =   345
         Index           =   10
         Left            =   1485
         TabIndex        =   29
         Top             =   2220
         Width           =   1155
      End
      Begin VB.TextBox txtField 
         Height          =   315
         Index           =   3
         Left            =   1485
         TabIndex        =   28
         Top             =   1395
         Width           =   3570
      End
      Begin VB.TextBox txtField 
         Height          =   345
         Index           =   2
         Left            =   1485
         TabIndex        =   27
         Top             =   945
         Width           =   2490
      End
      Begin VB.TextBox txtField 
         Height          =   345
         Index           =   1
         Left            =   1485
         TabIndex        =   26
         Top             =   510
         Width           =   2490
      End
      Begin VB.TextBox txtField 
         Height          =   345
         Index           =   11
         Left            =   1485
         TabIndex        =   25
         Top             =   2670
         Width           =   1170
      End
      Begin VB.ComboBox cboReportsTo 
         Height          =   315
         Left            =   1485
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1815
         Width           =   3570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   195
         Index           =   2
         Left            =   810
         TabIndex        =   37
         Top             =   1455
         Width           =   360
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   36
         Top             =   1035
         Width           =   810
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   35
         Top             =   585
         Width           =   825
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hire Date:"
         Height          =   195
         Index           =   11
         Left            =   435
         TabIndex        =   34
         Top             =   2295
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "Extension:"
         Height          =   195
         Index           =   12
         Left            =   405
         TabIndex        =   33
         Top             =   2745
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "Reports To:"
         Height          =   195
         Index           =   13
         Left            =   315
         TabIndex        =   32
         Top             =   1875
         Width           =   855
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID:"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   31
         Top             =   195
         Width           =   960
      End
      Begin VB.Label lblEmployeeID 
         AutoSize        =   -1  'True
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
         Left            =   1500
         TabIndex        =   30
         Top             =   165
         Width           =   60
      End
   End
   Begin MSComctlLib.TabStrip tabEmployees 
      Height          =   4140
      Left            =   75
      TabIndex        =   2
      Top             =   600
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   7303
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Company Info"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Personal Info"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmEmployees.frx":0173
      Top             =   30
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   6
      X2              =   476
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   6
      X2              =   476
      Y1              =   329
      Y2              =   329
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000003&
      Caption         =   "Employees "
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
      Left            =   -30
      TabIndex        =   3
      Top             =   0
      Width           =   7275
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_conn As Connection
Dim mrstEmployees As Recordset
Dim mvarBookmark As Variant
Dim mbIsNew As Boolean
Public Key As String

Private Const fldFirstName = 1
Private Const fldLastName = 2
Private Const fldTitle = 3
Private Const fldAddress = 4
Private Const fldCity = 5
Private Const fldRegion = 6
Private Const fldPostalCode = 7
Private Const fldHomePhone = 8
Private Const fldBirthDate = 9
Private Const fldHireDate = 10
Private Const fldExtension = 11
Private Const fldNotes = 12
Private Const fldReportsTo = 13
Private Const fldCountry = 14
Private Const fldTitleOfCourtesy = 15
Dim m_DataChanged(1 To 15) As Boolean
Public Sub EditRecord(cn As Connection, rs As Recordset)
Dim i As Long

On Error Resume Next
    Set m_conn = cn
    mvarBookmark = rs.Bookmark
    Set mrstEmployees = rs.Clone
    FillReportsToCombo
    FillCountryCombo
    mrstEmployees.Bookmark = mvarBookmark
    lblName = mrstEmployees![FirstName] & " " & mrstEmployees![LastName] & "   "
    lblEmployeeID = mrstEmployees![EmployeeID]
    txtField(fldFirstName).Text = mrstEmployees![FirstName]
    txtField(fldLastName).Text = mrstEmployees![LastName]
    txtField(fldTitle).Text = mrstEmployees![Title]
    txtField(fldAddress).Text = mrstEmployees![Address]
    txtField(fldCity).Text = mrstEmployees![City]
    txtField(fldRegion).Text = mrstEmployees![Region]
    txtField(fldPostalCode).Text = mrstEmployees![PostalCode]
    txtField(fldHomePhone).Text = mrstEmployees![Phone]
    txtField(fldBirthDate).Text = Format(mrstEmployees![BirthDate], "Medium Date")
    txtField(fldHireDate).Text = Format(mrstEmployees![HireDate], "Medium Date")
    txtField(fldExtension).Text = mrstEmployees![Extension]
    txtField(fldNotes).Text = mrstEmployees![Notes]
    cboCountry.Text = mrstEmployees![Country]
    cboTitleOfCourtesy.Text = mrstEmployees![TitleOfCourtesy]
    If IsNull(mrstEmployees![ReportsTo]) Then
        cboReportsTo.ListIndex = 0
    Else
        For i = 1 To cboReportsTo.ListCount - 1
            If cboReportsTo.ItemData(i) = mrstEmployees![ReportsTo] Then
                cboReportsTo.ListIndex = i
                Exit For
            End If
        Next
    End If
    Caption = "Employees - " & mrstEmployees![FirstName] & " " & mrstEmployees![LastName]
    For i = 1 To UBound(m_DataChanged)
        m_DataChanged(i) = False
    Next
    Me.Show
End Sub
Public Sub NewRecord(cn As Connection, rs As Recordset)
Dim i As Long

On Error Resume Next
    Set m_conn = cn
    Set mrstEmployees = rs.Clone
    mbIsNew = True
    mvarBookmark = Null
    FillReportsToCombo
    FillCountryCombo
    lblName = "New Employee  "
    cboReportsTo.ListIndex = 0
    Caption = "Employees - New Employee"
    Me.Show
End Sub


Private Sub cboCountry_Click()

    m_DataChanged(fldCountry) = True
    
End Sub

Private Sub cboCountry_Change()

    m_DataChanged(fldCountry) = True
    
End Sub

Private Sub cboReportsTo_Click()
    m_DataChanged(fldReportsTo) = True
    
End Sub

Private Sub cboTitleOfCourtesy_Click()
    m_DataChanged(fldTitleOfCourtesy) = True
End Sub

Private Sub cboTitleOfCourtesy_Change()
    m_DataChanged(fldTitleOfCourtesy) = True
End Sub

Private Sub cmdCancel_Click()

    Unload Me
End Sub


Private Sub cmdOK_Click()
On Error GoTo EH_cmdOK
Dim bUpdate As Boolean
Dim i As Integer

    For i = 1 To UBound(m_DataChanged)
        If m_DataChanged(i) Then bUpdate = True
    Next
    If bUpdate Then
        If mbIsNew Then
            mrstEmployees.AddNew
        End If
        If m_DataChanged(fldFirstName) Then mrstEmployees![FirstName] = txtField(fldFirstName)
        If m_DataChanged(fldLastName) Then mrstEmployees![LastName] = txtField(fldLastName)
        If m_DataChanged(fldTitle) Then mrstEmployees![Title] = txtField(fldTitle)
        If m_DataChanged(fldHireDate) Then mrstEmployees![HireDate] = TextToNull(txtField(fldHireDate))
        If m_DataChanged(fldExtension) Then mrstEmployees![Extension] = txtField(fldExtension)
        If m_DataChanged(fldAddress) Then mrstEmployees![Address] = txtField(fldAddress)
        If m_DataChanged(fldCity) Then mrstEmployees![City] = txtField(fldCity)
        If m_DataChanged(fldRegion) Then mrstEmployees![Region] = txtField(fldRegion)
        If m_DataChanged(fldPostalCode) Then mrstEmployees![PostalCode] = txtField(fldPostalCode)
        If m_DataChanged(fldHomePhone) Then mrstEmployees![HomePhone] = txtField(fldHomePhone)
        If m_DataChanged(fldBirthDate) Then mrstEmployees![BirthDate] = TextToNull(txtField(fldBirthDate))
        If m_DataChanged(fldNotes) Then mrstEmployees![Notes] = txtField(fldNotes)
        If m_DataChanged(fldCountry) Then mrstEmployees![Country] = cboCountry.Text
        If m_DataChanged(fldTitleOfCourtesy) Then mrstEmployees![TitleOfCourtesy] = cboTitleOfCourtesy.Text
        If m_DataChanged(fldReportsTo) Then
            If cboReportsTo.Text = "" Then
                mrstEmployees![ReportsTo] = Null
            Else
                mrstEmployees![ReportsTo] = cboReportsTo.ItemData(cboReportsTo.ListIndex)
            End If
        End If
        mrstEmployees.Update
        Hide
        frmMain.OnRecordUpdate CatalogEmployees, mvarBookmark
    End If
    Unload Me
    Exit Sub
    
EH_cmdOK:
    MsgBox Err.Description

End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmMain.UnloadForm Key

End Sub


Private Sub TabEmployees_BeforeClick(Cancel As Integer)

    picFrame(tabEmployees.SelectedItem.Index).Visible = False
End Sub



Private Sub tabEmployees_Click()

    picFrame(tabEmployees.SelectedItem.Index).Visible = True
    
End Sub

Private Sub txtField_Change(Index As Integer)

    m_DataChanged(Index) = True
    
End Sub


Private Sub FillCountryCombo()
Dim rstCountry As Recordset
Set rstCountry = New Recordset
    rstCountry.Open "SELECT DISTINCT Employees.Country FROM Employees ORDER BY Employees.Country", m_conn, adOpenForwardOnly, adLockReadOnly
    cboCountry.Clear
    cboCountry.AddItem ""
    Do Until rstCountry.EOF
        cboCountry.AddItem rstCountry![Country]
        rstCountry.MoveNext
    Loop

End Sub

Private Sub FillReportsToCombo()

    cboReportsTo.Clear
    cboReportsTo.AddItem ""
    Do Until mrstEmployees.EOF
        cboReportsTo.AddItem mrstEmployees![FirstName] & " " & mrstEmployees![LastName]
        cboReportsTo.ItemData(cboReportsTo.NewIndex) = mrstEmployees![EmployeeID]
        mrstEmployees.MoveNext
    Loop

End Sub
