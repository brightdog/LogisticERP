VERSION 5.00
Begin VB.UserControl UC 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   ScaleHeight     =   2220
   ScaleWidth      =   5415
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   900
      Top             =   1260
   End
   Begin VB.ListBox lstIncremental 
      Height          =   1860
      Left            =   3840
      TabIndex        =   4
      Top             =   300
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Text            =   "Address"
      Top             =   420
      Width           =   4455
   End
   Begin VB.TextBox txtDistrict 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Text            =   "District"
      Top             =   0
      Width           =   1635
   End
   Begin VB.TextBox txtCity 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1740
      TabIndex        =   1
      Text            =   "City"
      Top             =   0
      Width           =   1635
   End
   Begin VB.TextBox txtProvince 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Province"
      Top             =   0
      Width           =   1635
   End
End
Attribute VB_Name = "UC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mLastIncrementalControl As VB.Control
Dim mdicCity As New Scripting.Dictionary '3级连动菜单中，承上启下的，省份是独立的，不需要纪录，当前城市列表必须要保存，否则当前行政区就无法取得。
Dim bolArrowKeyDownState As Boolean '在用方向键选择增量下拉框里内容的时候，会出发onclick事件，坑爹啊！只好再加个窗体变量控制一下了。
Public gdicLocation As Scripting.Dictionary
Dim bolOnLoading As Boolean

Private Sub lstIncremental_Click()

    If lstIncremental.Text <> "" And (Not bolArrowKeyDownState) Then
        mLastIncrementalControl.Text = lstIncremental.Text
        lstIncremental.Visible = False
    End If

End Sub

Private Sub lstIncremental_KeyDown(KeyCode As Integer, Shift As Integer)

    If lstIncremental.Text <> "" And (Not bolArrowKeyDownState) Then
        mLastIncrementalControl.Text = lstIncremental.Text
        lstIncremental.Visible = False
    End If

End Sub

Private Sub Timer1_Timer()

    If Not ActiveControl Is Nothing Then
        Debug.Print ActiveControl.Name
        

    End If

End Sub

Private Sub txtAddress_GotFocus()
    lstIncremental.Visible = False
End Sub

Private Sub txtProvince_Change()

    If Not bolOnLoading Then
        If txtProvince.Text <> "" Then
            Call txtProvince_GotFocus
            txtCity.Text = ""
            txtDistrict.Text = ""
            'mdicCity.RemoveAll
        Else
            lstIncremental.Visible = False
        End If
    End If

End Sub

Private Sub txtProvince_GotFocus()
Debug.Print "txtProvince_GotFocus"
    lstIncremental.Visible = False
    Set mLastIncrementalControl = txtProvince
    lstIncremental.Visible = QuichSearchbyLocaldic(txtProvince, lstIncremental, gdicLocation.Item("0"))

End Sub

Private Sub txtProvince_KeyDown(KeyCode As Integer, Shift As Integer)
    bolArrowKeyDownState = True

    If lstIncremental.Visible Then

        Select Case True

            Case KeyCode = vbKeyDown

                If lstIncremental.ListIndex < lstIncremental.ListCount - 1 Then
                    lstIncremental.ListIndex = lstIncremental.ListIndex + 1
                End If

            Case KeyCode = vbKeyUp

                If lstIncremental.ListIndex > 0 Then
                    lstIncremental.ListIndex = lstIncremental.ListIndex - 1
                End If

            Case KeyCode = vbKeyReturn Or KeyCode = vbKeyTab

                If lstIncremental.ListIndex >= 0 Then
                    txtProvince.Text = lstIncremental.Text
                    txtCity.Text = ""
                    txtDistrict.Text = ""
                    mdicCity.RemoveAll
                    lstIncremental.Visible = False
                End If

                bolArrowKeyDownState = False
        End Select

    End If
    
End Sub
'===========================================================================

Private Sub txtCity_Change()

    If Not bolOnLoading Then
        If txtProvince.Text <> "" And (Not bolArrowKeyDownState) Then
            Call txtCity_GotFocus
        Else
            lstIncremental.Visible = False
        End If
    End If

End Sub

Private Sub txtCity_GotFocus()
    lstIncremental.Visible = False
    Set mLastIncrementalControl = txtCity
    Dim dic As Scripting.Dictionary
    Dim strFatherKey As String
    strFatherKey = FindDicKeyByValue(txtProvince.Text, gdicLocation.Item("0"))
    If strFatherKey <> "" Then
    Set dic = FindSubAreaByFather(strFatherKey)
    Set mdicCity = dic
    lstIncremental.Visible = QuichSearchbyLocaldic(txtCity, lstIncremental, dic)
    End If
End Sub

Private Sub txtCity_KeyDown(KeyCode As Integer, Shift As Integer)

    bolArrowKeyDownState = True

    If lstIncremental.Visible Then

        Select Case True

            Case KeyCode = vbKeyDown

                If lstIncremental.ListIndex < lstIncremental.ListCount - 1 Then
                    lstIncremental.ListIndex = lstIncremental.ListIndex + 1
                End If

            Case KeyCode = vbKeyUp

                If lstIncremental.ListIndex > 0 Then
                    lstIncremental.ListIndex = lstIncremental.ListIndex - 1
                End If

            Case KeyCode = vbKeyReturn Or KeyCode = vbKeyTab

                If lstIncremental.ListIndex >= 0 Then
                    txtCity.Text = lstIncremental.Text
                    txtDistrict.Text = ""
                    lstIncremental.Visible = False
                End If

        End Select

    End If

    bolArrowKeyDownState = False
End Sub

'===========================================================================
Private Sub txtDistrict_Change()

    If Not bolOnLoading Then
        If txtDistrict.Text <> "" And (Not bolArrowKeyDownState) Then
            Call txtDistrict_GotFocus
        Else
            lstIncremental.Visible = False
        End If
    End If

End Sub

Private Sub txtDistrict_GotFocus()
    lstIncremental.Visible = False
    Set mLastIncrementalControl = txtDistrict
    Dim dic As Scripting.Dictionary
    Dim strFatherKey As String
    strFatherKey = FindDicKeyByValue(txtCity.Text, mdicCity)

    If strFatherKey <> "" Then
        Set dic = FindSubAreaByFather(strFatherKey)
        lstIncremental.Visible = QuichSearchbyLocaldic(txtDistrict, lstIncremental, dic)
    End If

End Sub

Private Sub txtDistrict_KeyDown(KeyCode As Integer, Shift As Integer)
    bolArrowKeyDownState = True

    If lstIncremental.Visible Then

        Select Case True

            Case KeyCode = vbKeyDown

                If lstIncremental.ListIndex < lstIncremental.ListCount - 1 Then
                    lstIncremental.ListIndex = lstIncremental.ListIndex + 1
                End If

            Case KeyCode = vbKeyUp

                If lstIncremental.ListIndex > 0 Then
                    lstIncremental.ListIndex = lstIncremental.ListIndex - 1
                End If

            Case KeyCode = vbKeyReturn Or KeyCode = vbKeyTab

                If lstIncremental.ListIndex >= 0 Then
                    txtDistrict.Text = lstIncremental.Text
                    lstIncremental.Visible = False
                End If

        End Select

    End If

    bolArrowKeyDownState = False
End Sub

Private Sub UserControl_Initialize()
    bolOnLoading = True
    Dim iFile As Integer

    iFile = VBA.FreeFile()
    '==============2014-08-27 ==================================
    '目前为不判断版本，每次都直接从服务器取一份3级城市数据信息。
    '为兼容今后的版本控制，及载入速度，先加密存放到本地。
    '客户以及订单输入的时候，需要用到。
    '解密比加密快哈~~所以服务器端存放的是加密的版本，然后下载到本地存储密文后，解密直接使用。
    '加密慢，是因为拼接字符串引起的！！！换成StringBuilder之后秒开啊！！
    '不过服务器上为了安全考虑，存放密文比较好。
    '为了速度考虑，还是存放明文吧，一个90KB，一个200KB。。。
    '总结:加密也是要有代价的！
    '==========================================================
    On Error Resume Next
    Dim strFileContent As String
    Debug.Print Timer
    Open App.Path & "\..\Config\Location.Config" For Input As #iFile
    Input #iFile, strFileContent
    Close #iFile
    
    Dim objde As New clsDE
    Debug.Print Timer
    Set gdicLocation = JSON.Parse(objde.Decode(strFileContent))
    Debug.Print ":" & Timer
    'Debug.Print objDE.Decode(objDE.Encode(strHtml))
    Set objde = Nothing
    
    Dim ctl As VB.Control

    For Each ctl In Controls
    
        If TypeName(ctl) = "TextBox" Then
        
            ctl.Text = ""
        
        End If
    
    Next

    bolOnLoading = False
End Sub

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
                    If Not ctlFather Is Nothing Then
                        OffSetTop = OffSetTop + ctlFather.Top
124                     offSetLeft = offSetLeft + ctlFather.Left

                    

120                     Do While Not TypeName(ctlFather.Container) = ctlFather.Container.Name
                            Set ctlFather = ctlFather.Container
                            OffSetTop = OffSetTop + ctlFather.Top
                            offSetLeft = offSetLeft + ctlFather.Left

                        Loop

                    End If
                End If
                
128             ctlList.Move .Left + offSetLeft, .Top + OffSetTop + .Height, .width, 2000
            
            End With

            Dim i As Integer
            
130         For i = 0 To ctlList.ListCount - 1
                If Trim(ctl.Text) <> "" Then
132             If InStr(1, ctlList.List(i), ctl.Text, vbBinaryCompare) > 0 Then
134                 ctlList.Selected(i) = True
                    Exit For
                End If
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

Public Function FindDicKeyByValue(ByVal txt As String, ByRef dic As Scripting.Dictionary) As String

    Dim v As Variant
    FindDicKeyByValue = ""

    For Each v In dic.keys
    
        If dic.Item(v) = txt Then
            FindDicKeyByValue = CStr(v)
            Exit For
        End If
    
    Next

End Function

Public Function FindSubAreaByFather(ByVal strFatherKey As String) As Scripting.Dictionary

    Dim dic As Scripting.Dictionary
    
    Dim v As Variant
    
    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    
    Reg.Global = False
    Reg.IgnoreCase = False
    Reg.MultiLine = True
    Reg.Pattern = "," & strFatherKey & "$"
    
    For Each v In gdicLocation.keys

        If Reg.Test(CStr(v)) Then
        
            Set dic = gdicLocation.Item(v)
            Exit For
        End If
        
    Next
    
    Set FindSubAreaByFather = dic
End Function

Public Sub SetValueByJson(ByVal strJson As String)
    
    bolOnLoading = True
    Call CheckJsonisValued(strJson)
    Dim dic As Scripting.Dictionary
    Set dic = JSON.Parse(strJson)
    
    Dim v As Variant

    For Each v In dic.keys
    
        Call CallByName(Controls.Item("txt" & CStr(v)), "text", VbLet, dic.Item(v))
    
    Next
    
    bolOnLoading = False

End Sub

Private Function CheckJsonisValued(ByVal strJson As String) As Boolean
    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    
    Reg.Global = False
    Reg.IgnoreCase = True
    Reg.MultiLine = False
    
    Reg.Pattern = "\{""Province""\:""(.*?)"",""City""\:""(.*?)"",""District""\:""(.*?)"",""Address""\:""(.*?)""\}"

    If Reg.Test(strJson) Then
    
        CheckJsonisValued = True
    Else
        CheckJsonisValued = False
    
    End If
    
End Function
