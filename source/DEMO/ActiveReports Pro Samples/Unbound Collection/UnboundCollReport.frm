VERSION 5.00
Begin VB.Form frmCollectionUnboundReport 
   Caption         =   "Unbound Report"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPageCountInpageHeader 
      Caption         =   "Page Count in Page Header"
      Height          =   240
      Left            =   150
      TabIndex        =   5
      Top             =   750
      Width           =   2565
   End
   Begin VB.CheckBox chkLastSingle 
      Caption         =   "Last Group has a single record"
      Height          =   315
      Left            =   150
      TabIndex        =   4
      Top             =   1050
      Width           =   2565
   End
   Begin VB.CheckBox chkPageCount 
      Caption         =   "Page Count"
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   375
      Width           =   2565
   End
   Begin VB.CheckBox chkGrpKeepTogether 
      Caption         =   "GrpKeepTogether = All"
      Height          =   240
      Left            =   150
      TabIndex        =   2
      Top             =   75
      Width           =   2565
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   150
      TabIndex        =   1
      Top             =   2025
      Width           =   2490
   End
   Begin VB.CommandButton btnShowReport 
      Caption         =   "Show Report"
      Height          =   465
      Left            =   150
      TabIndex        =   0
      Top             =   1500
      Width           =   2490
   End
End
Attribute VB_Name = "frmCollectionUnboundReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnExit_Click()
    Unload Me
    End
End Sub

Private Sub btnShowReport_Click()
Dim rpt As New rptUnboundCollectionGrouping

    If chkGrpKeepTogether.Value = 1 Then
        rpt.ghOrder.GrpKeepTogether = ddGrpAll
    End If
    If chkPageCount.Value = 1 Then
        rpt.txtPage.SummaryRunning = ddSRAll
        rpt.txtPage.SummaryType = ddSMPageCount
        rpt.lblOf.Visible = True
        rpt.txtPageCount.SummaryType = ddSMPageCount
    End If
    If chkPageCountInpageHeader.Value = 1 Then
        rpt.txtHdrPage.SummaryType = ddSMPageCount
        rpt.txtHdrPage.SummaryRunning = ddSRAll
        rpt.lblHdrOf.Visible = True
        rpt.txtHdrPageCount.SummaryType = ddSMPageCount
    End If
    If chkLastSingle.Value = 1 Then
        rpt.bLastIsSingle = True
    End If
    rpt.InitCollection
    rpt.Show
End Sub

