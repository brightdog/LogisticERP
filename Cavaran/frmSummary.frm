VERSION 5.00
Begin VB.Form frmSummary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Summary"
   ClientHeight    =   3090
   ClientLeft      =   1620
   ClientTop       =   2235
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   2235
      TabIndex        =   1
      Top             =   2655
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Description"
      Height          =   2475
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   5760
      Begin VB.CommandButton cmdFormat 
         Caption         =   "Format..."
         Height          =   345
         Left            =   165
         TabIndex        =   5
         Top             =   1980
         Width           =   1245
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "Sort..."
         Height          =   345
         Left            =   165
         TabIndex        =   4
         Top             =   1455
         Width           =   1245
      End
      Begin VB.CommandButton cmdGroup 
         Caption         =   "Group By..."
         Height          =   345
         Left            =   165
         TabIndex        =   3
         Top             =   930
         Width           =   1245
      End
      Begin VB.CommandButton cmdFields 
         Caption         =   "Fields..."
         Height          =   345
         Left            =   165
         TabIndex        =   2
         Top             =   405
         Width           =   1245
      End
      Begin VB.Label lblSort 
         BackStyle       =   0  'Transparent
         Height          =   390
         Left            =   1500
         TabIndex        =   9
         Top             =   1455
         Width           =   4170
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblGroups 
         BackStyle       =   0  'Transparent
         Height          =   390
         Left            =   1500
         TabIndex        =   8
         Top             =   930
         Width           =   4170
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFields 
         BackStyle       =   0  'Transparent
         Height          =   390
         Left            =   1500
         TabIndex        =   7
         Top             =   405
         Width           =   4170
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblformat 
         Caption         =   "Fonts and other Table View settings"
         Height          =   390
         Left            =   1500
         TabIndex        =   6
         Top             =   1980
         Width           =   4170
      End
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_GridEX As GridEX



Private Sub cmdFields_Click()
    If frmShowfields.ShowFields(m_GridEX) Then
        LoadFieldNames
    End If
End Sub

Private Sub cmdFormat_Click()
    If m_GridEX.View = jgexTable Then
        frmTableview.FormatGrid m_GridEX
    Else
        frmCardView.FormatGrid m_GridEX
    End If
End Sub

Private Sub cmdGroup_Click()
    frmGroupBy.GroupGrid m_GridEX
    LoadGroupNames
End Sub

Private Sub cmdOK_Click()

    Hide
End Sub

Private Sub cmdSort_Click()
    frmSort.SortGrid m_GridEX
    LoadSortNames
End Sub

Public Sub ShowSummary(gr As GridEX)
Dim strTemp As String

    Set m_GridEX = gr
    LoadFieldNames
    LoadGroupNames
    LoadSortNames
    strTemp = "Fonts and other "
    If gr.View = jgexCard Then
        strTemp = strTemp & "Card"
        cmdGroup.Enabled = False
    Else
        strTemp = strTemp & "Table"
    End If
    strTemp = strTemp & " View settings"
    lblformat = strTemp
    Show 1
    Set m_GridEX = Nothing
    Unload Me
End Sub

Private Sub LoadFieldNames()
Dim strTemp As String
Dim c As JSColumn

    For Each c In m_GridEX.Columns
        If c.Visible Then
            strTemp = strTemp & c.Tag & ", "
        End If
    Next
    strTemp = Left(strTemp, Len(strTemp) - 2)
    lblFields = strTemp

End Sub

Private Sub LoadGroupNames()
Dim strTemp As String
Dim c As JSColumn
Dim gr As JSGroup

    For Each gr In m_GridEX.Groups
        Set c = m_GridEX.Columns(gr.ColIndex)
        strTemp = strTemp & c.Tag
        If gr.SortOrder = jgexSortAscending Then
            strTemp = strTemp & " (ascending), "
        Else
            strTemp = strTemp & " (descending), "
        End If
    Next
    If strTemp = "" Then
        strTemp = "None"
    Else
        strTemp = Left(strTemp, Len(strTemp) - 2)
    End If
    lblGroups = strTemp

End Sub


Private Sub LoadSortNames()
Dim strTemp As String
Dim c As JSColumn
Dim sk As JSSortKey

    For Each sk In m_GridEX.SortKeys
        Set c = m_GridEX.Columns(sk.ColIndex)
        strTemp = strTemp & c.Tag
        If sk.SortOrder = jgexSortAscending Then
            strTemp = strTemp & " (ascending), "
        Else
            strTemp = strTemp & " (descending), "
        End If
    Next
    If strTemp = "" Then
        strTemp = "None"
    Else
        strTemp = Left(strTemp, Len(strTemp) - 2)
    End If
    lblSort = strTemp

End Sub

