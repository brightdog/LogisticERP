VERSION 5.00
Begin VB.Form frmcfgServer 
   Caption         =   "Server Info"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmcfgServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   2460
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox txtHost 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   960
      TabIndex        =   1
      Top             =   660
      Width           =   3315
   End
   Begin VB.Label Label1 
      Caption         =   "Port:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1140
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Host:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmcfgServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Call modFormConfig.SaveConfig(Me)
End Sub

Private Sub Form_Load()
    Call modFormConfig.ReadConfig(Me)
    
End Sub
