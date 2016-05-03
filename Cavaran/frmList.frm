VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.0#0"; "GridEX20.ocx"
Begin VB.Form frmList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   1365
   ClientTop       =   2865
   ClientWidth     =   6660
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   5415
      TabIndex        =   2
      Top             =   735
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
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
      Left            =   5415
      TabIndex        =   1
      Top             =   360
      Width           =   1200
   End
   Begin GridEX20.GridEX gexList 
      Height          =   2550
      Left            =   75
      TabIndex        =   0
      Top             =   390
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   4498
      Version         =   "2.0"
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxInfoText=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   3
      ImagePicture1   =   "frmList.frx":000C
      ImagePicture2   =   "frmList.frx":011E
      ImagePicture3   =   "frmList.frx":0438
      DataMode        =   1
      GridLines       =   0
      BackColorBkg    =   -2147483624
      ColumnHeaderHeight=   285
      IntProp8        =   0   'False
      ColumnsCount    =   3
      Column(1)       =   "frmList.frx":0752
      Column(2)       =   "frmList.frx":08A2
      Column(3)       =   "frmList.frx":0982
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmList.frx":0A66
      FormatStyle(2)  =   "frmList.frx":0B9E
      FormatStyle(3)  =   "frmList.frx":0C4E
      FormatStyle(4)  =   "frmList.frx":0D02
      FormatStyle(5)  =   "frmList.frx":0DDA
      ImageCount      =   3
      ImagePicture(1) =   "frmList.frx":0E92
      ImagePicture(2) =   "frmList.frx":0FA4
      ImagePicture(3) =   "frmList.frx":12BE
      PrinterProperties=   "frmList.frx":15D8
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   45
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_OK As Boolean
Public Function ChooseSupplier(ByVal ConnectionString As String, ByVal SupplierID As String) As Variant
Dim cols As JSColumns
Dim col As JSColumn
Dim rsTemp As Recordset

    m_OK = False
    Set cols = gexList.Columns
    Set col = cols("Icon")
    col.DefaultIcon = 3
    Set col = cols("ID")
    col.Caption = "Supplier ID"
    col.DataField = "SupplierID"
    col.TextAlignment = jgexAlignRight
    col.SortType = jgexSortTypeNumeric
    Set col = cols("Name")
    col.Caption = "Supplier"
    col.DataField = "CompanyName"
    gexList.DatabaseName = ConnectionString
    gexList.RecordSource = "Suppliers"
    gexList.RecordsetType = jgexRSADOStatic
    gexList.HoldFields
    gexList.Rebind
    gexList.SortKeys.Add 3, jgexSortAscending
    Set Me.Icon = gexList.GridImages(3).Picture
    Me.Caption = "Suppliers List"
    lblInfo = "Select a supplier from the list:"
    gexList.RefreshSort
    Set rsTemp = gexList.ADORecordset
    rsTemp.MoveFirst
    rsTemp.Find "SupplierID=" & SupplierID
    If rsTemp.EOF Then
        gexList.Row = 0
    Else
        gexList.MoveToBookmark rsTemp.Bookmark
    End If
    Show 1
    If m_OK Then
        If gexList.Row <> 0 Then
            ChooseSupplier = gexList.Value(2)
        Else
            ChooseSupplier = ""
        End If
    Else
        ChooseSupplier = Null
    End If
    Unload Me
End Function

Public Function ChooseCustomer(ByVal ConnectionString As String, ByVal CustomerID As String) As Variant
Dim cols As JSColumns
Dim col As JSColumn
Dim rsTemp As Recordset

    m_OK = False
    Set cols = gexList.Columns
    Set col = cols("Icon")
    col.DefaultIcon = 3
    Set col = cols("ID")
    col.Caption = "Customer ID"
    col.DataField = "CustomerID"
    Set col = cols("Name")
    col.Caption = "Customer"
    col.DataField = "CompanyName"
    gexList.DatabaseName = ConnectionString
    gexList.RecordSource = "Customers"
    gexList.RecordsetType = jgexRSADOStatic
    gexList.HoldFields
    gexList.Rebind
    gexList.SortKeys.Add 3, jgexSortAscending
    Set Me.Icon = gexList.GridImages(2).Picture
    Me.Caption = "Customer List"
    lblInfo = "Select a customer from the list:"
    gexList.RefreshSort
    Set rsTemp = gexList.ADORecordset
    rsTemp.MoveFirst
    rsTemp.Find "CustomerID ='" & CustomerID & "'"
    If rsTemp.EOF Then
        gexList.Row = 0
    Else
        gexList.MoveToBookmark rsTemp.Bookmark
    End If
    Show 1
    If m_OK Then
        If gexList.Row <> 0 Then
            ChooseCustomer = gexList.Value(2)
        Else
            ChooseCustomer = ""
        End If
    Else
        ChooseCustomer = Null
    End If
    Unload Me
End Function


Private Sub cmdCancel_Click()

    Hide
    
End Sub

Private Sub cmdOK_Click()

    m_OK = True
    Hide
End Sub



Private Sub gexList_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)

    If Column.Index <> 1 Then
        If Column.SortOrder = jgexSortNone Then
            gexList.SortKeys.Clear
            gexList.SortKeys.Add Column.Index, jgexSortAscending
        Else
            If gexList.SortKeys(1).SortOrder = jgexSortAscending Then
                gexList.SortKeys(1).SortOrder = jgexSortDescending
            Else
                gexList.SortKeys(1).SortOrder = jgexSortAscending
            End If
        End If
    End If
End Sub

Private Sub gexList_DblClick()

    cmdOK_Click
    
End Sub

