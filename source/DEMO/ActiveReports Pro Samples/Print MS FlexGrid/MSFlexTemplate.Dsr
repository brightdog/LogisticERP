VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptMSFlexTemplate 
   Caption         =   "MS FlexGrid  Report Template"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   16245
   _ExtentY        =   14367
   SectionData     =   "MSFlexTemplate.dsx":0000
End
Attribute VB_Name = "rptMSFlexTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_grid As MSFlexGrid
Private bDone As Boolean

Property Set Grid(grd As MSFlexGrid)
Dim ctl As Object
Dim iLeft As Integer
Dim i As Integer
    Set m_grid = grd
    
    For i = 0 To m_grid.Cols - 1
        Set ctl = Detail.Controls.Add("DDActiveReports2.Field")
        ctl.Left = iLeft
        ctl.Top = 0
        ctl.Width = m_grid.ColWidth(i)
        ctl.Tag = i
        Fields.Add ctl.Name
        ctl.DataField = ctl.Name
        
        iLeft = iLeft + ctl.Width + 144
        PrintWidth = iLeft
    Next i
End Property

Private Sub ActiveReport_FetchData(eof As Boolean)
Static iRow As Integer
Dim ctl As Object
Dim i As Integer

    If iRow < m_grid.Rows Then
        For Each ctl In Detail.Controls
            m_grid.Row = iRow:  m_grid.Col = ctl.Tag
            Fields(ctl.Name).Value = m_grid.Text
        Next
        iRow = iRow + 1
        eof = False
    End If
End Sub
