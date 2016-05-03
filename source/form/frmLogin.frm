VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00D3D3D3&
   BorderStyle     =   0  'None
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":000C
   ScaleHeight     =   2865
   ScaleWidth      =   5985
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Cancel          =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   300
      Width           =   375
   End
   Begin VB.CommandButton cmdLogin 
      Appearance      =   0  'Flat
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   435
      Left            =   3600
      TabIndex        =   2
      Top             =   1980
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3D3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Password"
      Top             =   1500
      Width           =   3555
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3D3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "User Name"
      Top             =   780
      Width           =   3555
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub cmdLogin_Click()

    If CheckInputVaild(Me.txtUserName.Text, Me.txtPassword.Text) Then
        If CheckLogin(Me.txtUserName.Text, Me.txtPassword.Text) = "OK" Then
            gUSERNAME = Me.txtUserName.Text
            Me.Hide
            'frmSplash.Show
            frmMain.Show
            frmMain.tcMain.Visible = True
            If Not isOpen("frmOrder") Then
            
                'Set fOrder = New frmOrder
                Load frmOrder
                WriteLog "tcMain.InsertItem 0, ""订单系统"""
                frmMain.tcMain.InsertItem 0, "订单系统", frmOrder.hWnd, 0
                WriteLog "frmOrder.doSearch(0)"
                Call frmOrder.doSearch(0)
            Else
            
                Call frmMain.SetActiveTab("订单系统")
            End If
            Unload Me
        
        Else
        MsgBox "用户名或者密码错误", vbOKOnly + vbInformation, "提示信息"
        Me.txtUserName.SetFocus
        End If
    Else
    
        MsgBox "检查用户名或者密码的正确性", vbOKOnly + vbInformation, "提示信息"
        Me.txtUserName.SetFocus
    End If

End Sub

Private Function CheckInputVaild(ByVal strUserName As String, ByVal strPassword As String) As Boolean
    CheckInputVaild = True

    If Trim(strUserName) = "" Or Trim(Me.txtUserName.Text) = "User Name" Then
        CheckInputVaild = False
        Exit Function
    End If

    If Trim(strPassword) = "" Or Trim(Me.txtPassword.Text) = "Password" Then
        CheckInputVaild = False
        Exit Function
    End If

End Function



Private Sub Form_Load()
    Me.Show
    
    Me.txtUserName.SetFocus
End Sub

Private Sub txtUserName_GotFocus()
    If Trim(Me.txtUserName.Text) = "User Name" Then
    
        Call ChangeInputStyle(Me.txtUserName, "EDIT")
    
    End If
    
    Call txtPassword_GotFocus
End Sub

Private Sub txtUserName_LostFocus()
    If Trim(Me.txtUserName.Text) = "" Then
    
        Call ChangeInputStyle(Me.txtUserName, "DEMO", "User Name")
    
    End If
End Sub

Private Sub txtPassword_GotFocus()
    If Trim(Me.txtPassword.Text) = "Password" Then
    
        Call ChangeInputStyle(Me.txtPassword, "EDIT")
    Else
        Me.txtPassword.SelStart = 0
        Me.txtPassword.SelLength = Len(Me.txtPassword.Text)
    End If
End Sub

Private Sub txtPassword_LostFocus()
    If Trim(Me.txtPassword.Text) = "" Then
    
        Call ChangeInputStyle(Me.txtPassword, "DEMO", "Password")
    
    End If
End Sub
