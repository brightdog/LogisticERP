VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptProductList 
   Caption         =   "ProductList"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   21696
   _ExtentY        =   14023
   SectionData     =   "ProductList.dsx":0000
End
Attribute VB_Name = "rptProductList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private hFile As Integer

Private Sub ActiveReport_DataInitialize()
    hFile = FreeFile
    Open App.Path & "\Products.txt" For Input As #hFile
    
    ' This sets up the fields used in data binding
    Fields.Add "ProductID"
    Fields.Add "ProductName"
    Fields.Add "QuantityPerUnit"
    Fields.Add "UnitPrice"
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
Dim sLine As String
Dim arr() As String

    ' We reached the end of the file we exit leaving the
    ' eof parameter as True (default except on first call) that will
    ' tell AR that we are done feeding data
    ' otherwise we have to set the eof parameter to False so that
    ' AR continues fetching data, until we're done
    ' if the report had a data control, the value of the parameter
    ' will be ignored, AR will always follow the data control's recordset
    ' EOF property
    If VBA.eof(hFile) Then
        eof = True
        Exit Sub
    Else
        eof = False
    End If
    
    Line Input #hFile, sLine
    arr = Split(sLine, ",")
    
    ' Here we set the values of the fields that we defines as unbound
    ' or user defined.
    Fields("ProductID").Value = Val(arr(0))
    Fields("ProductName").Value = arr(1)
    Fields("QuantityPerUnit").Value = arr(4)
    Fields("UnitPrice").Value = arr(5)
End Sub

Private Sub ActiveReport_ReportEnd()
    If hFile <> 0 Then
        Close #hFile
    End If
End Sub

