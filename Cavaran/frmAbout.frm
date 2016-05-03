VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About The Janus GridEX - Advanced Sample"
   ClientHeight    =   3960
   ClientLeft      =   2910
   ClientTop       =   2520
   ClientWidth     =   5070
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2733.262
   ScaleMode       =   0  'User
   ScaleWidth      =   4760.993
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1575
      TabIndex        =   0
      Top             =   3405
      Width           =   1650
   End
   Begin VB.Image imglogo 
      Height          =   1815
      Left            =   3060
      Picture         =   "frmAbout.frx":000C
      Top             =   90
      Width           =   1935
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright© 1998 Janus Systems. All rights reserved."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   1320
      Width           =   2925
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "http://www.janusys.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   1
      Left            =   1035
      MouseIcon       =   "frmAbout.frx":12D7
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Chech the news about the Janus GridEX"
      Top             =   2460
      Width           =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "visit us at:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   5
      Top             =   2475
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "info@janusys.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   1050
      MouseIcon       =   "frmAbout.frx":1429
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "We want to hear from you..."
      Top             =   2760
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Janus Systems"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Top             =   2100
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "e-mail: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   2775
      Width           =   675
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Janus GridEX && Janus ButtonBar Advanced Sample"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   900
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   2325
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Private Const SW_SHOW = 5


Private Sub cmdOK_Click()
  Unload Me
End Sub


Private Sub Label1_Click(Index As Integer)
    Select Case Index
        Case 1
            Screen.MousePointer = vbArrowHourglass
            Call ShellExecute(Me.hWnd, "open", "http://www.janusys.com", vbNullString, CurDir$, SW_SHOW)
            Screen.MousePointer = vbNormal
        Case 4
            Screen.MousePointer = vbArrowHourglass
            Call ShellExecute(Me.hWnd, "open", "mailto:info@janusys.com", vbNullString, CurDir$, SW_SHOW)
            Screen.MousePointer = vbNormal
    End Select
End Sub


