VERSION 5.00
Begin VB.Form frmProductList 
   Caption         =   "Product List from Text File"
   ClientHeight    =   1170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   ScaleHeight     =   1170
   ScaleWidth      =   2865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnShow 
      Caption         =   "Show"
      Height          =   390
      Left            =   750
      TabIndex        =   0
      Top             =   150
      Width           =   1515
   End
End
Attribute VB_Name = "frmProductList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnShow_Click()
Dim rpt As New rptProductList
    rpt.Show
End Sub
