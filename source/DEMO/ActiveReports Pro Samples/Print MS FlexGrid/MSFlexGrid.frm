VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMSFlexGrid 
   Caption         =   "MS FlexGrid Print Sample"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnPreview 
      Caption         =   "&Preview Report"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Data dc 
      Caption         =   "Publishers"
      Connect         =   "Access"
      DatabaseName    =   "F:\Program Files\DevStudio\VB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT PubID, Name, Telephone, Fax FROM Publishers"
      Top             =   840
      Width           =   6495
   End
   Begin MSFlexGridLib.MSFlexGrid flxGrid 
      Bindings        =   "MSFlexGrid.frx":0000
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   10186
      _Version        =   393216
      FixedCols       =   0
      GridLines       =   0
      GridLinesFixed  =   3
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "frmMSFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnPreview_Click()
Dim rpt As New rptMSFlexTemplate

    Set rpt.Grid = flxGrid
    rpt.Show
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    flxGrid.Move 0, flxGrid.Top, ScaleWidth, ScaleHeight - flxGrid.Top
End Sub

Private Sub Form_Load()
    ' Set the database path to NorthWind in VB's installation direcory
    dc.DatabaseName = GetVBPath() & "\Biblio.MDB"
    flxGrid.Cols = 4
    flxGrid.ColWidth(0) = 1440
    flxGrid.ColWidth(1) = 2.5 * 1440
    flxGrid.ColWidth(2) = 1440
    flxGrid.ColWidth(3) = 1440
End Sub


