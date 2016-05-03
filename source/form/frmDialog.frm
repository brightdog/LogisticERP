VERSION 5.00
Begin VB.Form frmDialog 
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   Icon            =   "frmDialog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   5985
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   3420
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   1980
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   540
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   60
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   1
      Top             =   180
      Width           =   1275
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   180
      Width           =   4515
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const btnOK As Integer = 1
Const btnCancel As Integer = 2
Const btnSave As Integer = 4


Public msgTxt As String
Public msgTitle As String
Public msgBotton As Integer

Public Function ShowMsg() As String
    Me.Caption = msgTitle
    Me.txtMessage.Text = msgTxt
    
    Dim intBtnCnt As Integer
    intBtnCnt = 0
    
    If (msgBotton And btnOK) = btnOK Then
        Me.cmdOK.Visible = True
        intBtnCnt = intBtnCnt + 1
    Else
        Me.cmdOK.Visible = False
    End If
    
    If (msgBotton And btnCancel) = btnCancel Then
        Me.cmdCancel.Visible = True
        intBtnCnt = intBtnCnt + 1
    Else
        Me.cmdCancel.Visible = False
    End If
    
    If (msgBotton And btnSave) = btnSave Then
        Me.cmdSave.Visible = True
        intBtnCnt = intBtnCnt + 1
    Else
        Me.cmdSave.Visible = False
    End If
    
    Dim btnLeft As Long
    btnLeft = (Me.width - (intBtnCnt * (1215 + 200))) / intBtnCnt
    
    Dim btn As VB.Control
    
    For Each btn In Me.Controls
    
        If TypeName(btn) = "CommandButton" Then
            btn.Left = btnLeft * intBtnCnt
            intBtnCnt = intBtnCnt - 1
        End If
        
    Next
    
End Function

