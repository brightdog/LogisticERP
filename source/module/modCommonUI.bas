Attribute VB_Name = "modCommonUI"
Option Explicit
Public Const LEFT_MARGINTOP As Long = 1000
Public Const LEFT_MARGINBOTTOM As Long = 1000
Public Const LEFT_MARGINLEFT As Long = 100
Public Const LEFT_WIDTH As Long = 2000
Public Const SUBFORM_OFFSETLEFT As Long = 120
Public Const SUBFORM_OFFSETTOP As Long = -600

Public mLastIncrementalControl As VB.Control


Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4

Public Sub MouseClick(ByVal x As Long, ByVal y As Long)

    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, x, y, 0, 0
    'mouse_event MOUSEEVENTF_LEFTUP, x, y, 0, 0
End Sub

Public Sub ChangeInputStyle(ByRef txt As VB.TextBox, ByVal strType As String, Optional ByVal strDefault As String = "")

    Select Case strType
    
        Case "EDIT"
            
            txt.ForeColor = vbBlack
            txt.Text = ""
            
        Case "DEMO"
        
            txt.ForeColor = &H808080
            txt.Text = strDefault
        
    End Select

End Sub

Public Function InitLayout(ByRef frm As VB.Form) As String

    Dim strLayoutJson As String
    strLayoutJson = ReadLayoutConfig(frm.name)
    Dim dicLayoutJson As Scripting.Dictionary
    Set dicLayoutJson = JSON.Parse(strLayoutJson)
    
    Dim v As Variant
    
    For Each v In dicLayoutJson

        If Not IsEmpty(v) Then
            Dim x As Variant

            For Each x In dicLayoutJson.Item(v)
                Call CallByName(frm.Controls(CStr(v)), CStr(x), VbLet, dicLayoutJson.Item(v).Item(x))
        
            Next

        End If
    
    Next

End Function

Public Function InitTextBox(ByRef frm As VB.Form) As String

    Dim ctl As VB.Control

    For Each ctl In frm.Controls
    
        If TypeName(ctl) = "TextBox" Then
        
            ctl.Text = ""
        
        End If
    
    Next

End Function

Public Function ReadLayoutConfig(ByVal frmName As String) As String

    Dim strResult As String
    Dim Fso As Scripting.FileSystemObject
    Set Fso = New Scripting.FileSystemObject
    
    Dim Ts As Scripting.TextStream
    
    Set Ts = Fso.OpenTextFile(APP_CONFIG_PATH & frmName & ".Layout", ForReading, True, TristateFalse)

    If Not Ts.AtEndOfStream Then
        ReadLayoutConfig = Ts.ReadAll
    Else
    
        ReadLayoutConfig = "{}"
    
    End If

    Set Fso = Nothing

End Function

Public Function FillGrid(ByRef GRD As MSFlexGrid, ByRef dicList As Scripting.Dictionary) As String
    
    If Not dicList Is Nothing Then
    
        Dim i, j As Integer
    
        GRD.Clear
        GRD.rows = 1

        If TypeName(dicList.Item("Header")) = "Collection" Then
            GRD.Cols = dicList.Item("Header").Count

            For i = 1 To dicList.Item("Header").Count
    
                GRD.TextMatrix(0, i - 1) = modCommon.doMappingHeadTitle(dicList.Item("Header").Item(i))
    
            Next

            GRD.rows = dicList.Item("Rst").Count + 1

            For i = 1 To dicList.Item("Rst").Count
        
                For j = 1 To dicList.Item("Rst").Item(i).Count
        
                    GRD.TextMatrix(i, j - 1) = dicList.Item("Rst").Item(i).Item(j)
        
                Next
    
            Next
            
        End If

        GRD.Row = 0
    End If
    
End Function

Public Function FillFormTextBox(ByRef frm As VB.Form, ByVal dicJSON As Scripting.Dictionary) As String

    If Not dicJSON Is Nothing Then
        Dim i As Integer
    
        Dim dicFormControls As Scripting.Dictionary
        Set dicFormControls = GetAllControlsInForm(frm)
    
        For i = 1 To dicJSON.Item("Header").Count
            Debug.Print dicJSON.Item("Header").Item(i) & "**" & dicJSON.Item("Rst").Item(1).Item(i)

            If dicFormControls.Exists("txt" & dicJSON.Item("Header").Item(i)) Then
    
                Dim ctl As VB.Control
        
                Set ctl = frm.Controls("txt" & dicJSON.Item("Header").Item(i))

                If TypeName(ctl) = "TextBox" Then
                    Call CallByName(frm.Controls("txt" & dicJSON.Item("Header").Item(i)), "Text", VbLet, RestoreJsonField(dicJSON.Item("Rst").Item(1).Item(i)))
                    'Debug.Print "txt" & dicJson.Item("Rst").Item(1).Item(i)
                ElseIf TypeName(ctl) = "TDBDate" Or TypeName(ctl) = "TDBNumber" Then
                    Call CallByName(frm.Controls("txt" & dicJSON.Item("Header").Item(i)), "Value", VbLet, RestoreJsonField(dicJSON.Item("Rst").Item(1).Item(i)))
                
                End If

            ElseIf dicFormControls.Exists("cbo" & dicJSON.Item("Header").Item(i)) Then
            
                Set ctl = frm.Controls("cbo" & dicJSON.Item("Header").Item(i))
                
                If TypeName(ctl) = "ComboBox" Then

                    If ctl.Style = 2 Then '仅DropDownList无法设置text属性
                        
                        Dim j As Integer

                        For j = 0 To ctl.ListCount - 1

                            If ctl.List(j) = RestoreJsonField(dicJSON.Item("Rst").Item(1).Item(i)) Then
                
                                ctl.ListIndex = j
                                Exit For
                            End If

                        Next

                    Else
                        ctl.Text = RestoreJsonField(dicJSON.Item("Rst").Item(1).Item(i))
                    
                    End If
                End If

            Else
            
                Dim bolFindCtl As Boolean
                
                bolFindCtl = TrytoFindCtl(frm, "chk" & dicJSON.Item("Header").Item(i))
                
                If bolFindCtl Then
                    Set ctl = frm.Controls("chk" & dicJSON.Item("Header").Item(i))

                    If TypeName(ctl) = "CheckBox" Then
                        Call CallByName(frm.Controls("chk" & dicJSON.Item("Header").Item(i)), "value", VbLet, IIf(dicJSON.Item("Rst").Item(1).Item(i) = "True", 1, 0))

                    Else

                        '在窗体里没找到这个控件，但是数据库却返回了这个字段。
                        '说明应该是隐藏的值，该怎么处理呢？
                        '是不是窗体里需要设置一个模块级字典对象来存放这些信息呢？
                        '动态添加控件肯定是不现实的，容易引起意外。
                        '不过这样做的好处，就是数据库可以返回更多有用的信息，来辅助存放在窗体变量中。
                        '比如显示的时候是文本，数据库中存放的却是ID之类的。
                        '类似HTML中的<input type="hidden" name="xxx" value="yyy" />
                        '在窗体提交的过程中，别忘记把这个字典对象里的值也一起提交上去。
                        If frm.mdicHiddenFields.Exists("txt" & dicJSON.Item("Header").Item(i)) Then
                            frm.mdicHiddenFields.Item("txt" & dicJSON.Item("Header").Item(i)) = RestoreJsonField(dicJSON.Item("Rst").Item(1).Item(i))
                        Else
                            frm.mdicHiddenFields.Add ("txt" & dicJSON.Item("Header").Item(i)), RestoreJsonField(dicJSON.Item("Rst").Item(1).Item(i))
                        End If
                    End If

                Else

                    If frm.mdicHiddenFields.Exists(dicJSON.Item("Header").Item(i)) Then '为什么这里的key会重复呢？？2014-11-14
                    
                        frm.mdicHiddenFields.Item(dicJSON.Item("Header").Item(i)) = RestoreJsonField(dicJSON.Item("Rst").Item(1).Item(i))
                    Else
                        frm.mdicHiddenFields.Add dicJSON.Item("Header").Item(i), RestoreJsonField(dicJSON.Item("Rst").Item(1).Item(i))
                    End If
                End If
            End If

        Next

    Else
        '返回的数据不是JSON格式的。需要添加报错信息
    End If

End Function

Public Function TrytoFindCtl(ByRef frm As VB.Form, ByVal strName As String) As Boolean
    TrytoFindCtl = False
    Dim ctl As VB.Control
    
    For Each ctl In frm.Controls
    
        If ctl.name = strName Then
            
            TrytoFindCtl = True
            Exit For
        End If
    
    Next

End Function

Public Function GetAllControlsInForm(ByRef frm As VB.Form) As Scripting.Dictionary

    Dim ctl As VB.Control
    Dim dicControls As Scripting.Dictionary
    Set dicControls = New Scripting.Dictionary

    For Each ctl In frm.Controls
    
        If Not dicControls.Exists(ctl.name) Then
        
            dicControls.Add ctl.name, Null

        End If
    
    Next

    Set GetAllControlsInForm = dicControls

End Function

Public Function FillPageNavi(ByRef frm As VB.Form, ByRef dicList As Scripting.Dictionary) As String

    If Not dicList Is Nothing Then
        'picPagging
        'lblPageInfo
        'cmdPaggingFirst
        'cmdPaggingPrev
        'cmdPagging (0 - 9)
        'cmdPaggingNext
        'cmdPaggingLast
        'cboSkip
        '底层picPagging 的大小就不动了，当然能动是更好了。以免将来底色不统一
        'cmdPagging (0 - 9) 需要根据实际需要，是否显示，并且布局需要调整
        '翻页统一调用当前FORM的一个模块级变量 mPageNum 来纪录当前叶数，在考虑是否用cboSkip的下拉框来纪录？
        '最好是做成用户控件，暴露几个接口就完事了，但是时间上肯定是来不及的。
        '先就流氓做法，凑合着能用算了。
        Dim CurrentPage As String
        Dim PageCount As String
        Dim RsCount As String
        Dim PageSize As String
        Dim i As Integer
        Dim strPageStart As String
        Dim strPageEnd As String
        Dim indexj As Integer
    
        CurrentPage = dicList.Item("CurrentPage")
        PageCount = IIf(dicList.Item("PageCount") = "", 0, dicList.Item("PageCount"))
        PageSize = dicList.Item("PageSize")
        RsCount = IIf(dicList.Item("RsCount") = "", 0, dicList.Item("RsCount"))
    
        frm.lblPageInfo = CurrentPage & "/" & PageCount
        frm.cboSkip.Clear
    
        For i = 1 To PageCount
            frm.cboSkip.AddItem i

            If i = CurrentPage Then
                frm.cboSkip.ListIndex = i - 1
            End If

        Next

        frm.cmdPaggingFirst.Tag = 1
        frm.cmdPaggingPrev.Enabled = True
        frm.cmdPaggingNext.Enabled = True

        For i = 0 To 8
            frm.cmdPagging(i).BackColor = &H8000000F
            frm.cmdPagging(i).Caption = ""
            frm.cmdPagging(i).Visible = False
            frm.cmdPagging(i).Tag = ""
            'frm.cmdPagging(i).Enabled = True
        Next
    
        If PageCount <= 9 Then

            For i = 1 To PageCount

                frm.cmdPagging(i - 1).Visible = True
                frm.cmdPagging(i - 1).Tag = i
                frm.cmdPagging(i - 1).Caption = i

                If i = CurrentPage Then
                    
                    frm.cmdPagging(i - 1).BackColor = &H80000009
                    frm.cmdPagging(i - 1).Enabled = False
                Else
                    frm.cmdPagging(i - 1).Enabled = True
                End If
            
            Next

        ElseIf CurrentPage <= PageCount - 4 Then
        
            strPageStart = CurrentPage - 4
            

            If strPageStart < 1 Then

                strPageStart = 1
            
            End If
            
            strPageEnd = strPageStart + 8
            
            indexj = 0

            For i = strPageStart To strPageEnd
                
                If i >= 10 Then
                    frm.cmdPagging(indexj).width = 355
                Else
                    frm.cmdPagging(indexj).width = 255
                End If
                
                frm.cmdPagging(indexj).Visible = True
                frm.cmdPagging(indexj).Tag = i
                frm.cmdPagging(indexj).Caption = i

                If i = CurrentPage Then
                    
                    frm.cmdPagging(indexj).BackColor = &H80000009
                    frm.cmdPagging(indexj).Enabled = False
                Else
                    frm.cmdPagging(indexj).Enabled = True
                End If

                indexj = indexj + 1
            Next

        Else
            strPageStart = PageCount - 8
            indexj = 0
            
            For i = strPageStart To PageCount
            
                If i >= 10 Then
                    frm.cmdPagging(indexj).width = 355
                Else
                    frm.cmdPagging(indexj).width = 255
                End If
                
                frm.cmdPagging(indexj).Visible = True
                frm.cmdPagging(indexj).Tag = i
                frm.cmdPagging(indexj).Caption = i

                If i = CurrentPage Then

                    frm.cmdPagging(indexj).BackColor = &H80000009
                    frm.cmdPagging(indexj).Enabled = False
                Else
                    frm.cmdPagging(indexj).Enabled = True
                End If

                indexj = indexj + 1
            Next

        End If

        

        If CurrentPage = 1 Then
            frm.cmdPaggingLast.Enabled = True
            frm.cmdPaggingFirst.Enabled = False
            frm.cmdPaggingPrev.Enabled = False
            frm.cmdPaggingNext.Tag = CurrentPage + 1
            frm.cmdPaggingLast.Tag = PageCount
    
        ElseIf CurrentPage = PageCount Then
            frm.cmdPaggingNext.Enabled = False
            frm.cmdPaggingLast.Enabled = False
            frm.cmdPaggingPrev.Tag = CurrentPage - 1
            frm.cmdPaggingFirst.Enabled = True
        Else
            frm.cmdPaggingFirst.Enabled = True
            frm.cmdPaggingLast.Enabled = True
            frm.cmdPaggingPrev.Tag = CurrentPage - 1
            frm.cmdPaggingNext.Tag = CurrentPage + 1
            frm.cmdPaggingLast.Tag = PageCount
        
        End If

    End If

End Function

Public Sub CloseAllForms()
    Dim F As Form

    For Each F In Forms

        Unload F

    Next

End Sub

Public Function isOpen(ByVal frmName As String) As Boolean
    Dim F As Form

    For Each F In Forms

        If F.name = frmName Then
            isOpen = True
            Exit For
        End If

    Next

End Function

Public Function QuichSearchbyLocaldic(ByRef ctl As VB.Control, ByRef ctlList As VB.ListBox, ByRef dic As Scripting.Dictionary) As Boolean
        '<EhHeader>
        On Error GoTo QuichSearchbyLocaldic_Err
        '</EhHeader>

100     ctlList.Clear
        Dim intFindinList As Integer
102     intFindinList = -1
        Dim v As Variant

104     If dic.Count > 0 Then

106         For Each v In dic.keys
        
108             ctlList.AddItem CStr(dic.Item(v))
                
            Next

110         With ctl
                Dim OffSetTop, offSetLeft As Long
112             OffSetTop = 0
114             offSetLeft = 0

116             If TypeName(ctl.Container) <> "Form" Then
                    Dim ctlFather As VB.Control
118                 Set ctlFather = ctl.Container
                    OffSetTop = OffSetTop + ctlFather.Top
124                 offSetLeft = offSetLeft + ctlFather.Left

120                 Do While Not TypeName(ctlFather.Container) = ctlFather.Container.name
                        Set ctlFather = ctlFather.Container
                        OffSetTop = OffSetTop + ctlFather.Top
                        offSetLeft = offSetLeft + ctlFather.Left

                    Loop

                End If
                
128             ctlList.Move .Left + offSetLeft, .Top + OffSetTop + .Height, .width, 2000
            
            End With

            Dim i As Integer
            
130         For i = 0 To ctlList.ListCount - 1

132             If InStr(1, ctlList.List(i), ctl.Text, vbBinaryCompare) > 0 Then
134                 ctlList.Selected(i) = True
                    Exit For
                End If

            Next

136         QuichSearchbyLocaldic = True
        Else
138         QuichSearchbyLocaldic = False
        End If

        '<EhFooter>
        Exit Function

QuichSearchbyLocaldic_Err:
        QuichSearchbyLocaldic = False
        Resume Next
        '</EhFooter>
End Function

Public Function GetGridRowData(ByRef GRD As MSFlexGrid, ByVal RowNum As Integer) As String
    '没有原生的取整行的内容的函数，只能自己写一个玩玩了。

    Dim i As Integer
    Dim strResult As String

    For i = 0 To GRD.Cols - 1
    
        strResult = strResult & GRD.TextMatrix(RowNum, i) & vbTab
    
    Next

    GetGridRowData = strResult
End Function

Public Function LoadSampleListToGrdByID(ByVal TableName As String, ByVal QueryString As String, ByRef GRD As MSFlexGridLib.MSFlexGrid)
    Dim dicList As Scripting.Dictionary
    'Dim dicParam As Scripting.Dictionary
    'Set dicParam = makedicQueryParamfromQueryString(QueryString)
    
    Set dicList = SearchSimpleList("tblOrder", QueryString)
    
    Call FillGrid(GRD, dicList)
End Function

Public Function GetAllControlsPOI(ByRef frm As VB.Form) As Scripting.Dictionary
        '需要注意的是，该方法对于控件数组需要做额外的处理
        '必须将数组的下标也放在KEY里面，否则会出现重复的情况。
        '只存百分比，将来缩放的时候用
    
        'FUCK!!!无法判断当前控件是否是控件数组，似乎没有这个方法。。。狗日的，算了！只支持普通模式了。
        '<EhHeader>
        On Error GoTo GetAllControlsPOI_Err
        '</EhHeader>
    
        Dim h As Long
        Dim w As Long
        Dim dicResult As Scripting.Dictionary
100     Set dicResult = New Scripting.Dictionary
        
102     h = frm.Height
104     w = frm.width
    
        Dim ctl As VB.Control

106     For Each ctl In frm.Controls

            If Not dicResult.Exists(ctl.name) Then
108             dicResult.Add ctl, Array(ctl.Left / w, ctl.Top / h, ctl.width / w, ctl.Height / h)
            End If

        Next
        
110     Set GetAllControlsPOI = dicResult
        '<EhFooter>
        Exit Function

GetAllControlsPOI_Err:
        
        Resume Next
        
        Err.Clear
        '</EhFooter>
End Function

Public Sub ResizeFormControls(ByRef frm As VB.Form, ByVal dic As Scripting.Dictionary, Optional ByVal NeedZoomCTLSize As Boolean = True)

    On Error Resume Next
    
    Dim v As Variant
    Dim w As Long
    Dim h As Long
    
    w = frm.width
    h = frm.Height
    
    For Each v In dic.keys
    
        CallByName v, "Left", VbLet, dic.Item(v)(0) * w
        CallByName v, "Top", VbLet, dic.Item(v)(1) * h

        If NeedZoomCTLSize Then
            CallByName v, "Width", VbLet, dic.Item(v)(2) * w
            CallByName v, "Height", VbLet, dic.Item(v)(3) * h
        End If

    Next
    
End Sub

Public Function MakeSingleCheck(ByRef frm As VB.Form, ByRef chkCtl As VB.CheckBox)

    '根据当前CheckBox的Tag属性，判断其余各个Checkbox，如果Tag相同，并且Value = 1，则将value设置成0

    Dim ctl As VB.Control
    
    For Each ctl In frm.Controls

        If TypeName(ctl) = "CheckBox" Then
        
            'Debug.Print ctl.name & " <> " & chkCtl.name
            If ctl.name <> chkCtl.name Then
                If chkCtl.Value = 1 Then
                    If ctl.Tag = chkCtl.Tag Then
                        ctl.Value = 0
                    End If
                End If
        
            End If
        End If
        DoEvents
    Next

End Function

Public Function FillListBoxWithDic(ByRef Lst As VB.ListBox, ByRef dic As Scripting.Dictionary, ByVal fld As String) As String

    Lst.Clear

    If Not dic Is Nothing Then
        If dic.Exists("ERR") Then
            If dic.Item("ERR") = "NOT FOUND" Then
        
            Else
                MsgBox dic.Item("ERR")
            End If

            Exit Function
        End If

        Dim v As Variant
        Dim i As Integer
        Dim FldNO As Integer
        FldNO = 0

        For i = 1 To dic.Item("Header").Count
    
            If dic.Item("Header")(i) = fld Then
                FldNO = i
                Exit For
            End If
    
        Next
    
        For i = 1 To dic.Item("Rst").Count
    
            Lst.AddItem dic.Item("Rst")(i)(FldNO)
    
        Next

    End If

End Function

Public Function FillComboBoxWithDic(ByRef Cbo As VB.ComboBox, ByRef dic As Scripting.Dictionary, ByVal fld As String) As String

    Cbo.Clear

    If dic.Exists("ERR") Then
        If dic.Item("ERR") = "NOT FOUND" Then
        
        Else
            MsgBox dic.Item("ERR")
        End If

        Exit Function
    End If

    Dim v As Variant
    Dim i As Integer
    Dim FldNO As Integer
    FldNO = 0

    For i = 1 To dic.Item("Header").Count
    
        If dic.Item("Header")(i) = fld Then
            FldNO = i
            Exit For
        End If
    
    Next
    
    For i = 1 To dic.Item("Rst").Count
    
        Cbo.AddItem dic.Item("Rst")(i)(FldNO)
    
    Next

    '如果只有一条纪录的话，干脆直接选中算了，省得麻烦了。
    If Cbo.ListCount = 1 Then
        Cbo.ListIndex = 0
    End If
    
End Function

Public Function FillCboWithSampleDic(ByRef Cbo As VB.ComboBox, ByRef dic As Scripting.Dictionary)

    Dim v As Variant
    
    Cbo.Clear

    For Each v In dic.keys
    
        Cbo.AddItem dic.Item(v)
    
    Next

End Function

Public Sub ShowIncrementalSearchList(ByVal tblName As String, ByVal FldName As String, ByVal SearchType As String, _
   ByRef ctlInputBox As VB.Control, ByRef Lst As VB.ListBox)
    On Error GoTo iERR

    If ctlInputBox.Text <> "" Then
        Set mLastIncrementalControl = ctlInputBox
        Dim dicIncremental As Scripting.Dictionary
        
        Set dicIncremental = IncrementalSearch(tblName, FldName, SearchType, ctlInputBox.Text)
        Lst.Clear
        Dim i As Integer

        If dicIncremental.Item("Rst").Count > 0 Then

            For i = 1 To dicIncremental.Item("Rst").Count
        
                Lst.AddItem CStr(dicIncremental.Item("Rst").Item(i).Item(1))
    
            Next

            Lst.Move ctlInputBox.Left, ctlInputBox.Top + ctlInputBox.Height, ctlInputBox.width, 2000
            Lst.Selected(0) = True
            Lst.Visible = True
        Else
            Lst.Visible = False
            
        End If

    Else
        Lst.Visible = False
        Set mLastIncrementalControl = Nothing
    End If

    Exit Sub
    
iERR:
    Lst.Visible = False
End Sub

Public Sub SelectIncrementalResult(ByVal KeyCode As Integer, ByRef ctlInputBox As VB.Control, ByRef Lst As VB.ListBox)
    On Error GoTo iERR

    If Lst.Visible Then

        Select Case True

            Case KeyCode = vbKeyDown

                If Lst.ListIndex < Lst.ListCount - 1 Then
                    Lst.ListIndex = Lst.ListIndex + 1
                End If

            Case KeyCode = vbKeyUp

                If Lst.ListIndex > 0 Then
                    Lst.ListIndex = Lst.ListIndex - 1
                End If

            Case KeyCode = vbKeyReturn

                If Lst.ListIndex >= 0 Then
                    ctlInputBox.Text = Lst.Text
                    Lst.Visible = False
                End If

        End Select

    End If

    Exit Sub
iERR:
    Lst.Visible = False
End Sub

Public Sub SelectCurrentIncrementalText(ByRef Lst As VB.ListBox)

    On Error GoTo iERR

    If Lst.Text <> "" Then
    
        mLastIncrementalControl.Text = Lst.Text
        Lst.Visible = False
    End If

    Exit Sub
iERR:
    Lst.Visible = False
End Sub
    
Public Sub SetUILock(ByRef frm As VB.Form, ByVal LockState As Boolean)

    Dim ctl As VB.Control
    
    For Each ctl In frm.Controls
    
        If TypeName(ctl) = "TextBox" Or TypeName(ctl) = "ComboBox" Or TypeName(ctl) = "CheckBox" Then
        
            ctl.Enabled = LockState
        
        End If
    
    Next

End Sub
