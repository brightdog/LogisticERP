VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPrintPreview 
   Caption         =   "Print Preview"
   ClientHeight    =   7350
   ClientLeft      =   1875
   ClientTop       =   2430
   ClientWidth     =   8115
   Icon            =   "frmPrintPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   490
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   541
   Begin MSComctlLib.StatusBar stbPreview 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   7020
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GridEX20.GEXPreview GEXPreview1 
      Height          =   6885
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   12144
      BeginProperty ToolbarFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PageSetupText   =   "Page Set&up..."
      PrintText       =   "&Print..."
      CloseButtonText =   "&Close"
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
On Error Resume Next
    GEXPreview1.Move 0, 0, ScaleWidth, ScaleHeight - stbPreview.Height
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    frmMain.WindowState = Me.WindowState
    If Me.WindowState = 0 Then
        frmMain.Move Left, Top, Width, Height
    End If
    frmMain.Show
    Hide
    
End Sub


Private Sub GEXPreview1_BeforePaginating()
    stbPreview.SimpleText = "Paginating..."
    
End Sub

Private Sub GEXPreview1_OnCloseClick()

    Unload Me
 
End Sub

Private Sub GEXPreview1_PageChanged()
Dim bTwoPages As Boolean
Dim strPage As String
    With GEXPreview1
        If .Zoom = jgexZoomTwoPages Then
            If .CurrentPage < .TotalPages Then
                bTwoPages = True
            End If
        End If
        If bTwoPages Then
            strPage = .CurrentPage & " - " & .CurrentPage + 1
        Else
            strPage = .CurrentPage
        End If
        stbPreview.SimpleText = "Page " & strPage & " of " & .TotalPages
    End With
        
End Sub


