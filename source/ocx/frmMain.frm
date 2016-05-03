VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   5220
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdSetValue 
      Caption         =   "Set value"
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Top             =   1860
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   480
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1140
      Width           =   1395
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   2220
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1320
      Left            =   3840
      TabIndex        =   2
      Top             =   1020
      Width           =   1215
   End
   Begin CitySelect.UC UC1 
      Height          =   2235
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   5115
      _extentx        =   9022
      _extenty        =   1508
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSetValue_Click()
    UC1.SetValueByJson "{""Province"":""xxx"",""City"":""yyy"",""District"":""zzz"",""Address"":""abc123""}"
End Sub
