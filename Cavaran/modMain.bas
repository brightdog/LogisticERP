Attribute VB_Name = "modMain"
Option Explicit

Public Const CatalogCustomers = 1
Public Const CatalogSuppliers = 2
Public Const CatalogEmployees = 3
Public Const CatalogProducts = 4
Public Const CatalogOrders = 5


Public Const gListPageSize As Integer = 20

Public Sub Main()

    Screen.MousePointer = 11
    frmSplash.Show
    DoEvents
    Load frmMain
    DoEvents
    frmMain.Show
    Unload frmSplash
    Screen.MousePointer = 0
    
End Sub


Public Function TextToNull(strText As String) As Variant

    If strText = "" Then
        TextToNull = Null
    Else
        TextToNull = strText
    End If

End Function


Public Function FillGrid(ByRef Grd As GridEX20.GridEX, ByRef rst As ADODB.Recordset) As String

    Dim i, j As Long
    
    Grd.col
    
    Do While Not rst.EOF
    
        
    
        rst.MoveNext
    Loop



End Function
