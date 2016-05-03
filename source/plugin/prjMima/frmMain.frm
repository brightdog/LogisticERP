VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "密码加密解密"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3630
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpwd 
      Height          =   345
      Left            =   750
      TabIndex        =   3
      Top             =   870
      Width           =   2835
   End
   Begin VB.CommandButton Command2 
      Caption         =   "解密"
      Height          =   315
      Left            =   1890
      TabIndex        =   2
      Top             =   480
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加密"
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1125
   End
   Begin VB.TextBox txtstr 
      Height          =   345
      Left            =   750
      TabIndex        =   0
      Top             =   60
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "密文"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   930
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "明文"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    If Me.txtstr.Text <> "" Then
        Dim cls As New clsTransformPWD
        Me.txtpwd.Text = cls.TransFormPWD(Me.txtstr.Text)
    End If

End Sub

Private Sub Command2_Click()
    If Me.txtpwd.Text <> "" Then
        Dim cls As New clsTransformPWD
        Me.txtstr.Text = cls.deTransFormPWD(Me.txtpwd.Text)
    End If
End Sub


