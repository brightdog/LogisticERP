VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptPickupReceipt 
   Caption         =   "取件单"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10920
   StartUpPosition =   1  '所有者中心
   _ExtentX        =   19262
   _ExtentY        =   7752
   SectionData     =   "rptPickupReceipt.dsx":0000
End
Attribute VB_Name = "rptPickupReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dicData As Scripting.Dictionary

Dim iRow As Integer
Dim i As Integer


Private Sub ActiveReport_DataInitialize()
    Fields.Add "OrderID"
    Fields.Add "OrderCode"
    Fields.Add "SenderName"
    Fields.Add "SenderPhone"
    Fields.Add "SenderAddress"
    Fields.Add "PkgNum"
    Fields.Add "Remark"

    i = 1
    iRow = dicData.Count
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    If i > iRow Then
    
        EOF = True
        Exit Sub
    End If
    
    Fields("OrderID") = dicData.Item(i - 1).Item("OrderID")
    Fields("OrderCode") = dicData.Item(i - 1).Item("OrderCode")
    Fields("SenderName") = dicData.Item(i - 1).Item("SenderName")
    Fields("SenderPhone") = dicData.Item(i - 1).Item("SenderPhone")
    Fields("SenderAddress") = dicData.Item(i - 1).Item("SenderAddress")
    Fields("PkgNum") = dicData.Item(i - 1).Item("PkgNum")
    Fields("Remark") = dicData.Item(i - 1).Item("Remark")
    
    i = i + 1
    EOF = False

End Sub

