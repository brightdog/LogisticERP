VERSION 5.00
Begin VB.Form frmLeft 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   0  'None
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3360
   ControlBox      =   0   'False
   FillColor       =   &H80000004&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.FileListBox lstTree 
      Height          =   2430
      Left            =   300
      TabIndex        =   0
      Top             =   1560
      Width           =   2295
   End
End
Attribute VB_Name = "frmLeft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'    Me.Show
    'Call Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.lstTree.Top = 0
    Me.lstTree.Left = 0
    Me.lstTree.Width = Me.Width
    Me.lstTree.Height = Me.Height
End Sub
