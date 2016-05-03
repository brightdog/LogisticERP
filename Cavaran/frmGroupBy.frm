VERSION 5.00
Begin VB.Form frmGroupBy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Group By"
   ClientHeight    =   5145
   ClientLeft      =   4065
   ClientTop       =   2475
   ClientWidth     =   6015
   Icon            =   "frmGroupBy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   343
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboExpand 
      Height          =   315
      ItemData        =   "frmGroupBy.frx":014A
      Left            =   2130
      List            =   "frmGroupBy.frx":0155
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4665
      Width           =   2625
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   330
      Left            =   4740
      TabIndex        =   18
      Top             =   1020
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4740
      TabIndex        =   17
      Top             =   630
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   4740
      TabIndex        =   16
      Top             =   240
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Then by"
      Height          =   1005
      Index           =   3
      Left            =   480
      TabIndex        =   12
      Top             =   3555
      Width           =   4290
      Begin VB.OptionButton optOrder 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   7
         Left            =   2800
         TabIndex        =   15
         Top             =   555
         Width           =   1320
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Ascending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   2800
         TabIndex        =   14
         Top             =   225
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.ComboBox cboGroups 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   285
         Width           =   2595
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Then by"
      Height          =   1005
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   2420
      Width           =   4290
      Begin VB.OptionButton optOrder 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   2800
         TabIndex        =   11
         Top             =   555
         Width           =   1320
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Ascending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   2820
         TabIndex        =   10
         Top             =   225
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.ComboBox cboGroups 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   285
         Width           =   2595
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Then by"
      Height          =   1005
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1285
      Width           =   4290
      Begin VB.ComboBox cboGroups 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   285
         Width           =   2595
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Ascending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   2800
         TabIndex        =   6
         Top             =   225
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   2790
         TabIndex        =   5
         Top             =   555
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group Items by"
      Height          =   1005
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4290
      Begin VB.OptionButton optOrder 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   2800
         TabIndex        =   3
         Top             =   555
         Width           =   1320
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Ascending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   2800
         TabIndex        =   2
         Top             =   225
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.ComboBox cboGroups 
         Height          =   315
         Index           =   0
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   285
         Width           =   2595
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expand/Collapse:"
      Height          =   195
      Left            =   540
      TabIndex        =   20
      Top             =   4710
      Width           =   1260
   End
End
Attribute VB_Name = "frmGroupBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_OK As Boolean


Private Sub cboExpand_Change()

End Sub


Private Sub cboGroups_Click(Index As Integer)

    If cboGroups(Index).ListIndex = 0 Then
        optOrder(Index * 2).Enabled = False
        optOrder(Index * 2 + 1).Enabled = False
        optOrder(Index * 2).Value = True
        If Index < 3 Then
            cboGroups(Index + 1).ListIndex = 0
            cboGroups(Index + 1).Enabled = False
        End If
    Else
        optOrder(Index * 2).Enabled = True
        optOrder(Index * 2 + 1).Enabled = True
        optOrder(Index * 2).Value = True
        If Index < 3 Then
            cboGroups(Index + 1).Enabled = True
        End If
    End If
End Sub

Public Sub GroupGrid(grid As GridEX)
Dim cs As JSColumns
Dim c As JSColumn
Dim i As Integer
Dim j As Integer
Dim gr As JSGroup
Dim grs As JSGroups
    cboExpand.ListIndex = 0
    For j = 0 To 3
        cboGroups(j).AddItem "(None)"
    Next
    Set cs = grid.Columns
    For i = 1 To cs.Count
        Set c = cs.Item(i)
        If c.ColumnType <> jgexIcon Then
            For j = 0 To 3
                cboGroups(j).AddItem c.Tag
                cboGroups(j).ItemData(cboGroups(j).NewIndex) = c.Index
            Next
        End If
    Next i
    cboGroups(0).ListIndex = 0
    For i = 1 To grid.Groups.Count
        Set gr = grid.Groups.Item(i)
        cboGroups(i - 1).Text = cs.Item(gr.ColIndex).Caption
        If gr.SortOrder = 1 Then
            optOrder((i - 1) * 2).Value = True
        Else
            optOrder((i - 1) * 2 + 1).Value = True
        End If
    Next i
    m_OK = False
    Show 1
    Dim Changes As Boolean
    Set grs = grid.Groups
    If m_OK Then
        For i = 0 To 3
            If cboGroups(i).ListIndex = 0 Then
                If grs.Count > i Then
                    Changes = True
                    Exit For
                End If
            Else
                If grs.Count >= i + 1 Then
                    Set gr = grs.Item(i + 1)
                    If gr.ColIndex <> cboGroups(i).ItemData(cboGroups(i).ListIndex) Then
                        Changes = True
                        Exit For
                    Else
                        If optOrder(i * 2).Value Then
                            If gr.SortOrder <> 1 Then
                                Changes = True
                                Exit For
                            End If
                        Else
                            If gr.SortOrder <> -1 Then
                                Changes = True
                                Exit For
                            End If
                        End If
                    End If
                Else
                    Changes = True
                    Exit For
                End If
            End If
        Next
    End If
    If Changes Then
        Screen.MousePointer = 11
       ' grid.Parent.Refresh
        grs.Clear
        For i = 0 To 3
            If cboGroups(i).ListIndex <> 0 Then
                grs.Add cboGroups(i).ItemData(cboGroups(i).ListIndex), IIf(optOrder(i * 2), 1, -1)
            End If
        Next
        grid.RefreshGroups cboExpand.ItemData(cboExpand.ListIndex)
        Screen.MousePointer = 0
    End If
    Unload Me
End Sub

Private Sub cmdCancel_Click()

    Hide
    
End Sub


Private Sub cmdClearAll_Click()

    cboGroups(0).ListIndex = 0
End Sub


Private Sub cmdOK_Click()

    m_OK = True
    Hide
End Sub

