VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConfig 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSFlexGridLib.MSFlexGrid MSFG 
      Height          =   2115
      Left            =   600
      TabIndex        =   0
      Top             =   180
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3731
      _Version        =   393216
      AllowBigSelection=   0   'False
      Appearance      =   0
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub

Private Sub Form_Resize()
    Me.MSFG.Top = 100
    Me.MSFG.Left = 0
    Me.MSFG.Width = Me.Width
    Me.MSFG.Height = Me.Height = 500
End Sub
