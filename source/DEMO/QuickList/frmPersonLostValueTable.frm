VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPersonLostValueTable 
   Caption         =   "PersonLostValueTable"
   ClientHeight    =   7200
   ClientLeft      =   1560
   ClientTop       =   3105
   ClientWidth     =   13500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   13500
   Begin VB.ListBox tmpListHead 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      ItemData        =   "frmPersonLostValueTable.frx":0000
      Left            =   8460
      List            =   "frmPersonLostValueTable.frx":0022
      TabIndex        =   25
      Top             =   780
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "读取"
      Height          =   540
      Left            =   8340
      TabIndex        =   24
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer tmrSave 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   10500
      Top             =   1380
   End
   Begin VB.ListBox tmpList 
      Appearance      =   0  'Flat
      Height          =   1650
      ItemData        =   "frmPersonLostValueTable.frx":0044
      Left            =   8640
      List            =   "frmPersonLostValueTable.frx":0046
      TabIndex        =   22
      Top             =   780
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Remark"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   9
      Top             =   2700
      Width           =   3375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "添加(&A)"
      Height          =   540
      Left            =   6060
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "更新(&S)"
      Height          =   540
      Left            =   6060
      TabIndex        =   10
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(&D)"
      Height          =   540
      Left            =   7200
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      Height          =   540
      Left            =   7200
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&C)"
      Height          =   540
      Left            =   6000
      TabIndex        =   14
      Top             =   2340
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "IsCompany"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   8
      Top             =   2355
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Source"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   7
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Place"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   6
      Top             =   1725
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Date"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   1395
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Incident"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LostValue"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Top             =   765
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Address"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   420
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Name"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1035
   End
   Begin GridEX20.GridEX grd 
      Height          =   4095
      Left            =   120
      TabIndex        =   23
      Top             =   3060
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   7223
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ScrollToolTipColumn=   ""
      ColumnAutoResize=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      ColumnHeaderHeight=   270
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      IntProp8        =   -1  'True
      IntProp9        =   "frmPersonLostValueTable.frx":0048
      ColumnsCount    =   2
      Column(1)       =   "frmPersonLostValueTable.frx":00A2
      Column(2)       =   "frmPersonLostValueTable.frx":016A
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmPersonLostValueTable.frx":020E
      FormatStyle(2)  =   "frmPersonLostValueTable.frx":0352
      FormatStyle(3)  =   "frmPersonLostValueTable.frx":0402
      FormatStyle(4)  =   "frmPersonLostValueTable.frx":04B6
      FormatStyle(5)  =   "frmPersonLostValueTable.frx":058E
      ImageCount      =   0
      PrinterProperties=   "frmPersonLostValueTable.frx":0646
   End
   Begin VB.Label lblLabels 
      Caption         =   "Remark:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   2700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Source:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   2340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Place:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   2025
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Date:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   1695
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Incident:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LostValue:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   1065
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Address:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   735
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   1815
   End
End
Attribute VB_Name = "frmPersonLostValueTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rst As New Recordset
Dim bolAdd As Boolean
Dim ctlCurrent As Control
Dim RecordID As Integer
Dim intListTopCount As Integer
Dim intStep As Integer


Private Sub cmdRead_Click()
    Dim strTmp As String
    Dim tmp As String
    Dim ParamLine() As String
    Dim ParamCell() As String
    Dim i As Integer


    Open "Person.txt" For Input As #1

    Do While Not EOF(1)

        Line Input #1, tmp
        strTmp = strTmp & vbCrLf & tmp
    Loop

    Close #1

    Debug.Print strTmp

    ParamLine = Split(strTmp, vbCrLf, -1, vbTextCompare)

    For i = 0 To UBound(ParamLine)

        ParamCell = Split(ParamLine(i), "::", 2, vbTextCompare)
        If UBound(ParamCell) = 1 Then
            CallByName Me.txtFields(i), "text", VbLet, ParamCell(1)
        End If

    Next

End Sub

Private Sub Form_Click()
    CloseList
End Sub

Private Sub Form_Load()

    Rst.Open "select top 1 * from PersonLostValueTable order by id desc", Conn, adOpenStatic, adLockReadOnly

    grd.DatabaseName = "database.mdb"
    grd.RecordSource = "select * from PersonLostValueTable"

    RefreshGrid

    grd.MoveLast
    RecordID = grd.Row

    Dim i As Integer

    For i = 0 To 8

        Me.txtFields(i).Text = Rst.Fields(i) & ""

    Next

    Rst.Close
    
    If Not bolAdd Then
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    
    intListTopCount = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo AddErr

    If Not bolAdd Then

        Dim i As Integer

        For i = 0 To 8

            Me.txtFields(i).Text = ""

        Next

        bolAdd = True

        cmdUpdate.Caption = "保存(&S)"
        cmdAdd.Caption = "放弃"
        'cmdDelete.Enabled = False
        'cmdRefresh.Enabled = False

        grd.Enabled = False

        Me.txtFields(1).SetFocus

        tmrSave.Enabled = True

    Else

        cmdAdd.Caption = "添加(&A)"
        cmdUpdate.Caption = "更新(&S)"
        grd.Enabled = True
        
        Call Form_Load
        
        bolAdd = False
        
        tmrSave.Enabled = False

    End If

    Exit Sub
AddErr:
    MsgBox Err.Description
End Sub

Private Sub cmdAdd_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then

        Call cmdAdd_Click

    End If

End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo UpdateErr
    Dim strsql As String

    If bolAdd Then
    
        cmdUpdate.Enabled = False

        strsql = "insert into PersonLostValueTable (Name,Address,LostValue,Incident,[Date],Place,Source,Remark) values ( '" & Me.txtFields(1) & "','" & Me.txtFields(2) & "','" & Me.txtFields(3) & "','" & Me.txtFields(4) & "','" & Me.txtFields(5) & "','" & Me.txtFields(6) & "','" & Me.txtFields(7) & "','" & Me.txtFields(8) & "')"

        Debug.Print strsql

        Conn.BeginTrans
        Conn.Execute strsql
        Conn.CommitTrans

        UpdateDropDownList

        cmdUpdate.Caption = "更新(&S)"
        'cmdAdd.Enabled = True
        cmdAdd.Caption = "添加(&A)"
        'cmdDelete.Enabled = True
        'cmdRefresh.Enabled = True

        RefreshGrid

        DoEvents
        grd.MoveLast
        RecordID = grd.Value(1)
        RefreshTextBox RecordID

        grd.Enabled = True

        bolAdd = False

        tmrSave.Enabled = False
        
        cmdUpdate.Enabled = True

    Else

        If Me.txtFields(0).Text <> "" Then
            
            RecordID = grd.Row
            
            strsql = "update PersonLostValueTable set Name='" & Me.txtFields(1) & "',Address='" & Me.txtFields(2) & "',LostValue='" & Me.txtFields(3) & "',Incident='" & Me.txtFields(4) & "',[Date]='" & Me.txtFields(5) & "',Place='" & Me.txtFields(6) & "',Source='" & Me.txtFields(7) & "',Remark='" & Me.txtFields(8) & "' where [id] = " & Me.txtFields(0).Text
            Debug.Print strsql

            Conn.BeginTrans
            Conn.Execute strsql
            Conn.CommitTrans

            UpdateDropDownList

            RefreshGrid
            
            DoEvents

            grd.MoveLast
            
            DoEvents

            grd.Row = RecordID


        Else  '处理初始化的时候，忘记点添加按钮，就直接输入第一条纪录。

            strsql = "insert into PersonLostValueTable (Name,Address,LostValue,Incident,[Date],Place,Source,Remark) values ( '" & Me.txtFields(1) & "','" & Me.txtFields(2) & "','" & Me.txtFields(3) & "','" & Me.txtFields(4) & "','" & Me.txtFields(5) & "','" & Me.txtFields(6) & "','" & Me.txtFields(7) & "','" & Me.txtFields(8) & "')"

            Debug.Print strsql

            Conn.BeginTrans
            Conn.Execute strsql
            Conn.CommitTrans

            UpdateDropDownList

            cmdUpdate.Caption = "更新(&S)"

            cmdAdd.Caption = "添加(&A)"

            RefreshGrid

            DoEvents
            grd.MoveLast
            RecordID = grd.Value(1)
            RefreshTextBox RecordID

            grd.Enabled = True

            bolAdd = False

            tmrSave.Enabled = False

        End If

    End If

    cmdAdd.SetFocus

    Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub cmdUpdate_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then

        Call cmdUpdate_Click

    End If

End Sub



Private Sub cmdClose_Click()
    Unload Me
End Sub

'/////////////////////////////////////////

Private Sub ShowList()
    tmpList.Visible = True
    tmpListHead.Visible = True
End Sub

Private Sub CloseList()
    tmpList.Visible = False
    tmpListHead.Visible = False
End Sub

Private Sub RefreshGrid()
    grd.Rebind
    grd.Refresh
End Sub

Private Sub UpdateDropDownList()
    Dim txtControl As Control

    For Each txtControl In Me.txtFields

        If txtControl.Index <> 0 Then

            CheckandInsert txtControl, Me.Name

        End If

    Next

End Sub

Private Sub RefreshTextBox(ByVal RecordID As String)

    If RecordID > 0 Then

        Rst.Open "select * from PersonLostValueTable where [id] = " & RecordID, Conn, adOpenStatic, adLockReadOnly

        Dim i As Integer

        For i = 0 To 8

            Me.txtFields(i).Text = Rst.Fields(i) & ""

        Next

        Rst.Close

    End If

End Sub

Private Sub grd_AfterUpdate()
    grd.Update
    RefreshTextBox RecordID
End Sub

Private Sub grd_Click()
    RecordID = grd.GetRowData(grd.Row).Value(1)

    RefreshTextBox RecordID

End Sub

Private Sub tmpList_DblClick()
    FillText
    CloseList
End Sub

Private Sub tmpList_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case vbKeyReturn
            FillText

        Case vbKeyEscape

            CloseList
            ctlCurrent.SetFocus
            
        Case vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9
        
            FillText Chr(KeyAscii)
          
    End Select

End Sub

Private Sub tmpList_KeyUp(KeyCode As Integer, Shift As Integer)

    If tmpList.ListIndex = 0 And KeyCode = vbKeyUp Then

        If intListTopCount >= 2 Then
            ctlCurrent.SetFocus
            DoEvents
            ctlCurrent.SelStart = Len(ctlCurrent.Text)
            DoEvents
            DoEvents
            CloseList

            intListTopCount = 1
            
            Exit Sub

        End If
        intListTopCount = intListTopCount + 1

    End If

End Sub


Private Sub tmpList_LostFocus()
    CloseList
End Sub

Public Sub DropDownList(ByVal KeyCode As Integer)
    On Error Resume Next

    If KeyCode = vbKeyDown Then
        tmpList.SetFocus
        tmpList.Selected(0) = True
    End If

End Sub

Private Sub MoveList(ByVal ctl As Control)
    tmpList.Top = ctl.Top + ctl.Height
    tmpList.Left = ctl.Left + 200
    tmpList.Width = ctl.Width - 200

    tmpListHead.Top = ctl.Top + ctl.Height
    tmpListHead.Left = ctl.Left


    If tmpList.ListCount > 0 Then
        ShowList
    Else
        CloseList
    End If
End Sub

Public Sub QuickSearchLastUse(ByVal ctl As Control)
    Set ctlCurrent = ctl
    tmpList.Clear

    Dim rstLast As New Recordset
    Dim strsql As String

    strsql = "select text from dropdown where KeyName = '" & Name & "." & ctl.Name & ctl.Index & "' order by lastusetime desc"

    rstLast.Open strsql, Conn

    Do While Not rstLast.EOF

        tmpList.AddItem rstLast.Fields(0)
        rstLast.MoveNext

    Loop


    MoveList ctl


End Sub

Public Sub QuickSearchLikeName(ByVal ctl As Control)

    Set ctlCurrent = ctl

    tmpList.Clear

    Dim rstLast As New Recordset
    Dim strsql As String

    strsql = "select top 10 [text] from dropdown where KeyName = '" & Name & "." & ctl.Name & ctl.Index & "' and [text] like '" & ctl.Text & "%'"

    rstLast.Open strsql, Conn

    Do While Not rstLast.EOF

        tmpList.AddItem rstLast.Fields(0)
        rstLast.MoveNext

    Loop


    MoveList ctl


End Sub

Public Sub FillText(Optional ByVal ItemIndex As Integer = -1)

    If ItemIndex = -1 Then

        ItemIndex = tmpList.ListIndex
    Else
        ItemIndex = ItemIndex - 1

    End If

    If ItemIndex <= tmpList.ListCount Then

        ctlCurrent.Text = tmpList.List(ItemIndex)
        ctlCurrent.SetFocus


        CloseList
        SendKeys vbTab
        
    End If

End Sub

Private Sub tmrSave_Timer()
    Dim strTmp As String

    Dim i As Integer

    For i = 1 To Me.txtFields.Count - 1

        strTmp = strTmp & "txtFields(" & i & ")::" & Me.txtFields(i).Text & vbCrLf

    Next

    Open "Person.txt" For Output As #1

    Print #1, strTmp

    Close #1

    strTmp = ""
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Sub txtFields_GotFocus(Index As Integer)

    Select Case Index

        Case 4, 5, 6, 7, 8

            If Me.txtFields(Index).Text = "" Then
                QuickSearchLastUse Me.txtFields(Index)

                If tmpList.ListCount > 0 Then
                    Me.txtFields(Index).Text = tmpList.List(0)
                    QuickSearchLastUse Me.txtFields(Index)
                Else
                    CloseList
                End If

            End If
            

    End Select

    txtFields(Index).SelStart = 0
    txtFields(Index).SelLength = Len(txtFields(Index).Text)
End Sub

Private Sub txtFields_KeyDown(Index As Integer, _
            KeyCode As Integer, _
            Shift As Integer)

    Set ctlCurrent = Me.txtFields(Index)


    Select Case KeyCode

        Case vbKeyDown

            If Me.txtFields(Index).Text = "" Then
                QuickSearchLastUse txtFields(Index)
                tmpList.SetFocus
                tmpList.Selected(0) = True
            Else

                If tmpList.Visible = False Then
                    QuickSearchLikeName txtFields(Index)
                End If

                If tmpList.Visible = True Then

                    tmpList.SetFocus
                    tmpList.Selected(0) = True
                End If

            End If

        Case vbKeyReturn

            SendKeys vbTab
            CloseList

        Case vbKeyTab, vbKeyEscape

            CloseList



    End Select

End Sub

Private Sub txtFields_KeyUp(Index As Integer, _
                            KeyCode As Integer, _
                            Shift As Integer)

    If Me.txtFields(Index).Text <> "" And KeyCode <> vbKeyReturn And KeyCode <> vbKeyBack And KeyCode <> vbKeyDelete Then
        QuickSearchLikeName txtFields(Index)
        

    ElseIf KeyCode = vbKeyF1 Then
    
        QuickSearchLastUse txtFields(Index)
        
        If tmpList.Visible = True Then
            
            ctlCurrent.Text = tmpList.List(0)
        
        End If
        
    End If

End Sub

Private Sub txtFields_LostFocus(Index As Integer)

    If Index = 5 And IsNumeric(Me.txtFields(5).Text) Then

        If Left$(Me.txtFields(5), 2) <> "19" Then

            Me.txtFields(5).Text = "19" & Me.txtFields(5).Text

            Dim tmp As String
            tmp = Me.txtFields(5).Text

            Select Case Len(Me.txtFields(5))

                Case 8
                    tmp = Left$(tmp, 4) & "-" & Right$(tmp, 4)
                    tmp = Left$(tmp, 7) & "-" & Right$(tmp, 2)
                Case 7
                    tmp = Left$(tmp, 4) & "-0" & Right$(tmp, 3)
                    tmp = Left$(tmp, 7) & "-" & Right$(tmp, 2)
                Case 6
                    tmp = Left$(tmp, 4) & "-" & Right$(tmp, 2)

            End Select

            Me.txtFields(5).Text = tmp

        End If

    End If

End Sub

Private Sub txtFields_MouseDown(Index As Integer, _
            Button As Integer, _
            Shift As Integer, _
            X As Single, _
            Y As Single)
    QuickSearchLastUse txtFields(Index)

End Sub
