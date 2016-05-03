VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Bot_Run"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3405
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   3405
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.VScrollBar vsThreadNum 
      Height          =   315
      Left            =   2520
      Max             =   49
      TabIndex        =   5
      Top             =   60
      Value           =   40
      Width           =   195
   End
   Begin VB.TextBox txtThreadNum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "10"
      Top             =   60
      Width           =   375
   End
   Begin VB.ListBox lstExpressNO 
      Height          =   3660
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   2235
   End
   Begin VB.CommandButton cmdManualRun 
      Caption         =   "Run"
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   60
      Width           =   645
   End
   Begin VB.TextBox txtExpressNO 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8700
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "06:00:00|16:00:00"
      Top             =   420
      Width           =   3375
   End
   Begin VB.Label lblState 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bolStop As Boolean


Private Sub cmdManualRun_Click()

    If Me.cmdManualRun.Caption = "Run" Then
        Me.cmdManualRun.Caption = "Stop"
        bolStop = False
        Dim i As Integer
        i = 0
        
        Dim arr() As String
        
        arr = Split(Me.txtExpressNO.Text, vbCrLf, -1, vbBinaryCompare)
        
        Dim dicExpressNO As Scripting.Dictionary
        Set dicExpressNO = New Scripting.Dictionary
        
        For i = 0 To UBound(arr)
        
            If Not dicExpressNO.Exists(arr(i)) And arr(i) <> "" Then
            
                dicExpressNO.Add arr(i), ""
            
            End If
        
        Next

        Me.lstExpressNO.Clear
        Dim v As Variant
        
        For Each v In dicExpressNO.Keys
        
            Me.lstExpressNO.AddItem CStr(v)
        
        Next
        
        Me.txtExpressNO.Visible = False
        Me.lstExpressNO.Visible = True
        i = 0

        Do While i <= Me.lstExpressNO.ListCount - 1 And Not bolStop
        
            If modProc.GetProcessCountbyName("ExpressBot.exe") < CInt(Me.txtThreadNum.Text) Then
            
                Shell App.Path & "\ExpressBot.exe " & Me.lstExpressNO.List(i), vbMinimizedNoFocus
                Me.lstExpressNO.List(i) = "-->" & vbTab & Me.lstExpressNO.List(i)
                i = i + 1
            Else
            
                Do
                    MySleep 0.1
                Loop While modProc.GetProcessCountbyName("ExpressBot.exe") >= CInt(Me.txtThreadNum.Text)
            
            End If
            
            Me.lblState.Caption = i & "/" & Me.lstExpressNO.ListCount
            MySleep 0.01
        Loop
        
    Else
    
        Me.cmdManualRun.Caption = "Run"
        bolStop = True
    
    End If

    Me.txtExpressNO.Visible = True
    Me.lstExpressNO.Visible = False
        
End Sub




Private Sub Form_Load()

    Call Form_Resize
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
100     End
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    Me.txtExpressNO.Height = Me.Height - Me.lblState.Height - 450
    Me.txtExpressNO.Width = Me.Width - 230
    Me.lstExpressNO.Visible = False
    Me.lstExpressNO.Top = Me.txtExpressNO.Top
    Me.lstExpressNO.Left = Me.txtExpressNO.Left
    Me.lstExpressNO.Width = Me.txtExpressNO.Width
    Me.lstExpressNO.Height = Me.txtExpressNO.Height
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
100     End
End Sub



Private Sub vsThreadNum_Change()
    Me.txtThreadNum.Text = 50 - Me.vsThreadNum.value
End Sub
