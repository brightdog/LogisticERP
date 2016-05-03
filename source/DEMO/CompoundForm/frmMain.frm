VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "CompoundForm"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   10305
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command3 
      Caption         =   "Load Content"
      Height          =   435
      Left            =   3900
      TabIndex        =   2
      Top             =   2700
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Button Bar"
      Height          =   435
      Left            =   2040
      TabIndex        =   1
      Top             =   2700
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Left Tree"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   2700
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Const mintOffsetLeft As Integer = 120
Public mHeight As Long
Public mWidth As Long
Public mLeft As String

Private Sub Command1_Click()
    Load frmLeft
    SetParent frmLeft.hWnd, frmMain.hWnd
    frmLeft.Show
    Call movefrmLeft(mWidth, mHeight, mLeft)
End Sub

Private Sub movefrmLeft(mWidth, mHeight, mLeft)
    frmLeft.Width = 2000

    frmLeft.Height = mHeight - 1000 - 1000
    
    frmLeft.Left = -mLeft - mintOffsetLeft + 100
    frmLeft.Top = 1000
End Sub

Private Sub Command2_Click()
    SetParent frmHeader.hWnd, frmMain.hWnd
    frmHeader.Show
    Call movefrmHeader(mWidth, mHeight, mLeft)
End Sub

Private Sub movefrmHeader(mWidth, mHeight, mLeft)
    frmHeader.Width = Me.Width '- 3000 - 1000
    frmHeader.Left = -mLeft - mintOffsetLeft + 200
    frmHeader.Top = 100
End Sub

Private Sub Command3_Click()
    SetParent frmContent.hWnd, frmMain.hWnd
    frmContent.Show
    Call movefrmContent(mWidth, mHeight, mLeft)
End Sub

Private Sub movefrmContent(mWidth, mHeight, mLeft)
    
    frmContent.Width = mWidth - 2000

    frmContent.Height = mHeight - 1000 - 1000
    frmContent.Left = -mLeft - mintOffsetLeft + 2100
    frmContent.Top = 1000
End Sub

Private Sub Form_Load()
    Me.Show
    Call Form_Resize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If isOpen("frmLeft") Then
        Unload frmLeft
    End If
    If isOpen("frmHeader") Then
        Unload frmHeader
    End If
    If isOpen("frmContent") Then
        Unload frmContent
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Me.Command1.Top = Me.ScaleHeight - Me.Command1.Height - 20
    Me.Command2.Top = Me.Command1.Top
    Me.Command3.Top = Me.Command1.Top
    mHeight = Me.Height
    mWidth = Me.Width
    mLeft = Me.Left
    If isOpen("frmLeft") Then Call movefrmLeft(mWidth, mHeight, mLeft)
    If isOpen("frmHeader") Then Call movefrmHeader(mWidth, mHeight, mLeft)
    If isOpen("frmContent") Then Call movefrmContent(mWidth, mHeight, mLeft)
End Sub
Function isOpen(fName As String) As Boolean
    Dim f As Form

    For Each f In Forms

        If f.Name = fName Then
            isOpen = True
            Exit For
        End If

    Next

End Function
