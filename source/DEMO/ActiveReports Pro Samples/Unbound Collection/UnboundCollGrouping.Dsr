VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptUnboundCollectionGrouping 
   Caption         =   "Unbound Grouping and Totals"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10635
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   18759
   _ExtentY        =   13653
   SectionData     =   "UnboundCollGrouping.dsx":0000
End
Attribute VB_Name = "rptUnboundCollectionGrouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bLastIsSingle As Boolean

Private colOrders As New Orders

Private Sub ActiveReport_DataInitialize()
    Fields.Add "OrderID"
    Fields.Add "ProductID"
    Fields.Add "ProductName"
    Fields.Add "Qty"
    Fields.Add "Price"
    Fields.Add "Amount"
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
Static i As Integer
    i = i + 1
    If i > colOrders.Count Then
        eof = True
        Exit Sub
    Else
        eof = False
    End If
    
    Fields("OrderID").Value = colOrders(i).Key
    Fields("ProductID").Value = colOrders(i).ProductID
    Fields("ProductName").Value = colOrders(i).ProductName
    Fields("Qty").Value = colOrders(i).Qty
    Fields("Price").Value = colOrders(i).Price
    Fields("Amount").Value = colOrders(i).Qty * colOrders(i).Price
    
End Sub

Public Sub InitCollection()
Dim o As Order
    With colOrders
        Set o = .Add("100"): o.ProductID = 1: o.ProductName = "Hard Drive": o.Qty = 1: o.Price = 435
        Set o = .Add("100"): o.ProductID = 2: o.ProductName = "CD ROM": o.Qty = 1: o.Price = 199
        Set o = .Add("100"): o.ProductID = 3: o.ProductName = "32MB RAM": o.Qty = 2: o.Price = 85
        Set o = .Add("101"): o.ProductID = 1: o.ProductName = "Hard Drive": o.Qty = 1: o.Price = 435
        Set o = .Add("102"): o.ProductID = 4: o.ProductName = "PII Processor": o.Qty = 1: o.Price = 500
        Set o = .Add("102"): o.ProductID = 5: o.ProductName = "Graphics Card": o.Qty = 1: o.Price = 189
        Set o = .Add("102"): o.ProductID = 3: o.ProductName = "32MB RAM": o.Qty = 4: o.Price = 85
        Set o = .Add("103"): o.ProductID = 1: o.ProductName = "Hard Drive": o.Qty = 1: o.Price = 435
        Set o = .Add("103"): o.ProductID = 2: o.ProductName = "CD ROM": o.Qty = 2: o.Price = 199
        Set o = .Add("103"): o.ProductID = 3: o.ProductName = "32MB RAM": o.Qty = 6: o.Price = 85
        Set o = .Add("103"): o.ProductID = 4: o.ProductName = "PII Processor": o.Qty = 2: o.Price = 500
        Set o = .Add(IIf(bLastIsSingle, "104", "103")): o.ProductID = 5: o.ProductName = "Graphics Card": o.Qty = 2: o.Price = 189
    End With
    
End Sub
