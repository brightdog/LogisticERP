VERSION 5.00
Begin VB.Form frmShowfields 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Show Fields"
   ClientHeight    =   3810
   ClientLeft      =   2175
   ClientTop       =   2145
   ClientWidth     =   7710
   Icon            =   "frmShowfields.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   6555
      TabIndex        =   9
      Top             =   450
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6555
      TabIndex        =   8
      Top             =   840
      Width           =   1035
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Move Down"
      Height          =   345
      Left            =   5280
      TabIndex        =   7
      Top             =   3345
      Width           =   1170
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Move Up"
      Height          =   345
      Left            =   4020
      TabIndex        =   6
      Top             =   3360
      Width           =   1140
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<- Remove"
      Height          =   345
      Left            =   2745
      TabIndex        =   5
      Top             =   930
      Width           =   1125
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add ->"
      Height          =   345
      Left            =   2745
      TabIndex        =   4
      Top             =   480
      Width           =   1125
   End
   Begin VB.ListBox lstVisible 
      Height          =   2790
      Left            =   3990
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.ListBox lstAvail 
      Height          =   2790
      Left            =   180
      TabIndex        =   0
      Top             =   495
      Width           =   2415
   End
   Begin VB.Label lblcaption 
      Caption         =   "Show these fields in this order:"
      Height          =   255
      Index           =   1
      Left            =   4035
      TabIndex        =   3
      Top             =   180
      Width           =   2460
   End
   Begin VB.Label lblcaption 
      Caption         =   "Available fields:"
      Height          =   255
      Index           =   0
      Left            =   195
      TabIndex        =   2
      Top             =   225
      Width           =   1800
   End
End
Attribute VB_Name = "frmShowfields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim m_OK As Boolean

Public Function ShowFields(gr As GridEX) As Boolean
Dim c As JSColumn
Dim strName As String
Dim i As Integer

    m_OK = False

    For i = 1 To gr.Columns.Count
        Set c = gr.Columns.ItemByPosition(i)
        strName = c.Tag
        If Not c.Visible Then
            lstAvail.AddItem strName
            lstAvail.ItemData(lstAvail.NewIndex) = c.Index
        Else
            lstVisible.AddItem strName
            lstVisible.ItemData(lstVisible.NewIndex) = c.Index
        End If
    Next
    On Error Resume Next
    lstAvail.ListIndex = 0
    lstVisible.ListIndex = 0
    Show 1
    If m_OK Then
        ShowFields = True
        For i = 0 To lstAvail.ListCount - 1
            Set c = gr.Columns(lstAvail.ItemData(i))
            c.Visible = False
        Next
        For i = 0 To lstVisible.ListCount - 1
            Set c = gr.Columns(lstVisible.ItemData(i))
            c.Visible = True
            c.ColPosition = i + 1
        Next
    End If
    Unload Me
End Function

Private Sub cmdAdd_Click()
Dim ColIndex As Integer
Dim ColText As String
Dim lngListindex As Long

Dim c As JSColumn
    lstAvail.SetFocus
    If lstAvail.ListIndex = -1 Then Exit Sub
    lngListindex = lstAvail.ListIndex
    
    ColIndex = lstAvail.ItemData(lngListindex)
    ColText = lstAvail.Text
    lstAvail.RemoveItem lngListindex
    lstVisible.AddItem ColText
    lstVisible.ItemData(lstVisible.NewIndex) = ColIndex
    If lstAvail.ListCount - 1 >= lngListindex Then
        lstAvail.ListIndex = lngListindex
    Else
        lstAvail.ListIndex = lngListindex - 1
    End If
    lstVisible.ListIndex = lstVisible.NewIndex
    EnableButtons
    
End Sub


Private Sub EnableButtons()

    cmdAdd.Enabled = (lstAvail.ListIndex <> -1)
    cmdRemove.Enabled = (lstVisible.ListIndex <> -1)
    cmdUp.Enabled = (lstVisible.ListIndex > 0)
    cmdDown.Enabled = (lstVisible.ListIndex < lstVisible.ListCount - 1)
    
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdDown_Click()
Dim ColIndex As Long
Dim ColText As String
Dim lngListindex As Long


    If lstVisible.ListIndex = -1 Or lstVisible.ListIndex = lstVisible.ListCount - 1 Then Exit Sub
    With lstVisible
        lngListindex = .ListIndex
        ColText = .Text
        ColIndex = .ItemData(lngListindex)
        .RemoveItem lngListindex
        lngListindex = lngListindex + 1
        .AddItem ColText, lngListindex
        .ItemData(.NewIndex) = ColIndex
        .ListIndex = .NewIndex
        .SetFocus
    End With
    EnableButtons

End Sub

Private Sub cmdOK_Click()
    m_OK = True
    Hide
End Sub

Private Sub cmdRemove_Click()
Dim ColIndex As Integer
Dim ColText As String
Dim lngListindex As Long

Dim c As JSColumn
    lstVisible.SetFocus
    If lstVisible.ListIndex = -1 Then Exit Sub
    lngListindex = lstVisible.ListIndex
    
    ColIndex = lstVisible.ItemData(lngListindex)
    ColText = lstVisible.Text
    lstVisible.RemoveItem lngListindex
    lstAvail.AddItem ColText
    lstAvail.ItemData(lstAvail.NewIndex) = ColIndex
    If lstVisible.ListCount - 1 >= lngListindex Then
        lstVisible.ListIndex = lngListindex
    Else
        lstVisible.ListIndex = lngListindex - 1
    End If
    lstAvail.ListIndex = lstAvail.NewIndex
    EnableButtons
    
End Sub


Private Sub cmdUp_Click()
Dim ColIndex As Long
Dim ColText As String
Dim lngListindex As Long


    If lstVisible.ListIndex <= 0 Then Exit Sub
    With lstVisible
        lngListindex = .ListIndex
        ColText = .Text
        ColIndex = .ItemData(lngListindex)
        .RemoveItem lngListindex
        If lngListindex > 0 Then lngListindex = lngListindex - 1
        .AddItem ColText, lngListindex
        .ItemData(.NewIndex) = ColIndex
        .ListIndex = .NewIndex
        .SetFocus
    End With
    EnableButtons
    
End Sub

Private Sub lstAvail_DblClick()

    cmdAdd_Click
    
End Sub


Private Sub lstVisible_Click()
    EnableButtons
    
End Sub

Private Sub lstVisible_DblClick()
    cmdRemove_Click
End Sub


