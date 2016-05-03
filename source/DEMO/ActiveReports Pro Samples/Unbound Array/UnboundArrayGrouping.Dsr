VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptUnboundArrayGrouping 
   Caption         =   "Unbound Grouping and Totals"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10635
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   18759
   _ExtentY        =   13653
   SectionData     =   "UnboundArrayGrouping.dsx":0000
End
Attribute VB_Name = "rptUnboundArrayGrouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bLastIsSingle As Boolean

Private arr(1 To 12) As OrderItem

Private iRow As Integer
Private tmpAmount As Currency
Private bNewGroup As Boolean

Private Type OrderItem
    OrderNo As Long
    ProductID As Long
    ProductName As String
    Qty As Integer
    Price As Currency
    Amount As Currency
End Type

Private Sub ActiveReport_DataInitialize()
    Fields.Add "OrderID"
    Fields.Add "ProductID"
    Fields.Add "ProductName"
    Fields.Add "Qty"
    Fields.Add "Price"
    Fields.Add "Amount"
    iRow = LBound(arr)
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
    If iRow > UBound(arr) Then
        eof = True
        Exit Sub
    End If
    Fields("OrderID") = arr(iRow).OrderNo
    Fields("ProductID") = arr(iRow).ProductID
    Fields("ProductName") = arr(iRow).ProductName
    Fields("Qty") = arr(iRow).Qty
    Fields("Price") = arr(iRow).Price
    Fields("Amount") = arr(iRow).Amount
    
    ' It is important to set EOF to False if there are more records
    ' otherwise the default True value will be used and the report
    ' will stop
    eof = False
    iRow = iRow + 1
End Sub

Public Sub InitArray()
    arr(1).OrderNo = 100: arr(1).ProductID = 1: arr(1).ProductName = "Hard Drive": arr(1).Qty = 1: arr(1).Price = 435: arr(1).Amount = arr(1).Qty * arr(1).Price
    arr(2).OrderNo = 100: arr(2).ProductID = 2: arr(2).ProductName = "CD ROM": arr(2).Qty = 1: arr(2).Price = 199: arr(2).Amount = arr(2).Qty * arr(2).Price
    arr(3).OrderNo = 100: arr(3).ProductID = 3: arr(3).ProductName = "32MB RAM": arr(3).Qty = 2: arr(3).Price = 85: arr(3).Amount = arr(3).Qty * arr(3).Price
    arr(4).OrderNo = 101: arr(4).ProductID = 1: arr(4).ProductName = "Hard Drive": arr(4).Qty = 1: arr(4).Price = 435: arr(4).Amount = arr(4).Qty * arr(4).Price
    arr(5).OrderNo = 102: arr(5).ProductID = 4: arr(5).ProductName = "PII Processor": arr(5).Qty = 1: arr(5).Price = 500: arr(5).Amount = arr(5).Qty * arr(5).Price
    arr(6).OrderNo = 102: arr(6).ProductID = 5: arr(6).ProductName = "Graphics Card": arr(6).Qty = 1: arr(6).Price = 189: arr(6).Amount = arr(6).Qty * arr(6).Price
    arr(7).OrderNo = 102: arr(7).ProductID = 3: arr(7).ProductName = "32MB RAM": arr(7).Qty = 4: arr(7).Price = 85: arr(7).Amount = arr(7).Qty * arr(7).Price
    arr(8).OrderNo = 103: arr(8).ProductID = 1: arr(8).ProductName = "Hard Drive": arr(8).Qty = 1: arr(8).Price = 435: arr(8).Amount = arr(8).Qty * arr(8).Price
    arr(9).OrderNo = 103: arr(9).ProductID = 2: arr(9).ProductName = "CD ROM": arr(9).Qty = 2: arr(9).Price = 199: arr(9).Amount = arr(9).Qty * arr(9).Price
    arr(10).OrderNo = 103: arr(10).ProductID = 3: arr(10).ProductName = "32MB RAM": arr(10).Qty = 6: arr(10).Price = 85: arr(10).Amount = arr(10).Qty * arr(10).Price
    arr(11).OrderNo = 103: arr(11).ProductID = 4: arr(11).ProductName = "PII Processor": arr(11).Qty = 2: arr(11).Price = 500: arr(11).Amount = arr(11).Qty * arr(11).Price
    arr(12).OrderNo = IIf(bLastIsSingle, 104, 103): arr(12).ProductID = 5: arr(12).ProductName = "Graphics Card": arr(12).Qty = 2: arr(12).Price = 189: arr(12).Amount = arr(12).Qty * arr(12).Price
End Sub
