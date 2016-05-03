VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmConfig 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmConfig.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin TrueDBGrid60.TDBGrid TDBGrd 
      Height          =   1695
      Left            =   420
      OleObjectBlob   =   "frmConfig.frx":000C
      TabIndex        =   0
      Top             =   300
      Width           =   3255
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As XArrayDB

Private Sub cmdSave_Click()
    Dim obj As clsWebConfig
    Set obj = New clsWebConfig
    Call obj.SaveXml(ConvertGridtoArray(Me.TDBGrd))
    Set obj = Nothing
End Sub

Private Sub Form_Load()
    modAppPath.Init
    Dim obj As clsWebConfig
    Set obj = New clsWebConfig
    Set x = New XArrayDB
    x.LoadRows (obj.LoadXml)

    

    Dim row As Long, col As Integer

    ' The LowerBound and UpperBound properties correspond
    ' to the LBound and UBound functions in Visual Basic.
    ' Hard-coded dimensions can be used instead, if known.
'    For row = x.LowerBound(1) To x.UpperBound(1)
'        For col = x.LowerBound(2) To x.UpperBound(2)
'            x(row, col) = "Row " & row & ", Col " & col
'        Next col
'    Next row

    ' Bind True DBGrid Control to this XArrayDB instance

    Set TDBGrd.Array = x

End Sub

Private Function ConvertGridtoArray(ByRef ctl As TDBGrid) As String()

    Dim strResult() As String
    
    Dim i, j As Long
    
    ReDim strResult(ctl.record, x.LowerBound(2) To x.UpperBound(2))
    
    For i = LBound(strResult, 1) To UBound(strResult, 1)
    
        For j = LBound(strResult, 2) To UBound(strResult, 2)
        
            strResult(i, j) = x(i, j)
        
        Next
    
    Next

    ConvertArrayDB = strResult
End Function


Private Sub Form_Resize()
On Error Resume Next
    Me.TDBGrd.Top = 100
    Me.TDBGrd.Left = 0
    Me.TDBGrd.Width = Me.Width - 120
    Me.TDBGrd.Height = Me.Height - 1000
    Me.cmdSave.Top = Me.Height - 800
    Me.cmdSave.Left = (Me.Width / 2) - (Me.cmdSave.Width / 2)
End Sub

