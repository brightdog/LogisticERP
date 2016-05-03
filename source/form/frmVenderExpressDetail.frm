VERSION 5.00
Begin VB.Form frmVenderExpressDetail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "物流信息"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   Icon            =   "frmVenderExpressDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8535
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   495
      Left            =   3180
      TabIndex        =   1
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox txtExpressDetail 
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
      Height          =   4935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmVenderExpressDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mdicHiddenFields As New Scripting.Dictionary

Public Function LoadDetail(ByVal ExpressNOList As String) As String
    
    Dim strURL, strPostData As String
    strPostData = "{""Type"":""Load_ExpressDetail"",""ExpressNOList"":""" & ExpressNOList & """}"
    strURL = "plugin/venderexpress.asp"
    Dim strResult As String
    strResult = PostData(strURL, strPostData)
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)

    If IsObject(dicResult) Then
        If dicResult.Item("Rst").Count > 0 Then
    
            Dim SB As clsStringBuilder
            Set SB = New clsStringBuilder
            Dim i As Integer
            Dim intFldVenderDTPos As Integer
            Dim intFldVenderDescPos As Integer
            Dim intSiteCodePos As Integer
            Dim intExpressNOPos As Integer

            For i = 1 To dicResult.Item("Header").Count
    
                If dicResult.Item("Header").Item(i) = "VenderDT" Then
                    intFldVenderDTPos = i
                ElseIf dicResult.Item("Header").Item(i) = "VenderDesc" Then
                    intFldVenderDescPos = i
                ElseIf dicResult.Item("Header").Item(i) = "SiteCode" Then
                    intSiteCodePos = i
                ElseIf dicResult.Item("Header").Item(i) = "ExpressNO" Then
                    intExpressNOPos = i
                End If

            Next

            Dim strLastSiteCode As String
            Dim strLastExpressNO As String

            For i = 1 To dicResult.Item("Rst").Count
                Dim strSplitLine As String
                strSplitLine = "==============" & dicResult.Item("Rst").Item(i).Item(intSiteCodePos) & "_" & dicResult.Item("Rst").Item(i).Item(intExpressNOPos) & "===============" & vbCrLf
                If strLastSiteCode = "" Then
                    strLastSiteCode = dicResult.Item("Rst").Item(i).Item(intSiteCodePos)
                    strLastExpressNO = dicResult.Item("Rst").Item(i).Item(intExpressNOPos)
                    SB.Append strSplitLine
                Else

                    If strLastSiteCode <> dicResult.Item("Rst").Item(i).Item(intSiteCodePos) Or strLastExpressNO <> dicResult.Item("Rst").Item(i).Item(intExpressNOPos) Then
                        
                        SB.Append strSplitLine
                        
                        strLastSiteCode = dicResult.Item("Rst").Item(i).Item(intSiteCodePos)
                        strLastExpressNO = dicResult.Item("Rst").Item(i).Item(intExpressNOPos)
                    End If
                End If
            
                SB.Append dicResult.Item("Rst").Item(i).Item(intFldVenderDTPos) & vbTab & dicResult.Item("Rst").Item(i).Item(intFldVenderDescPos) & vbCrLf

            Next

            Me.txtExpressDetail.Text = SB.toString
            Set SB = Nothing
        Else
    
            Me.txtExpressDetail.Text = "暂无物流信息" & vbCrLf & ExpressNOList
    
        End If

    Else
        Unload Me
    End If

End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub


