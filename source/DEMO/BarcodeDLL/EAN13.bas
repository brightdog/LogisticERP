Attribute VB_Name = "EAN13and8"
'/****************************************************************************
' * Summary   : 条形码生成程序
' * Version   : 1.00
' * Start Date: 2004-6-07
' * My home   : http://www.mndsoft.com
' * E-Mail    : Mnd@Mndsoft.Com
' ****************************************************************************/

Dim LeftHand_Odd() As Variant
Dim LeftHand_Even() As Variant
Dim Right_Hand() As Variant
Dim Parity() As Variant

Dim BarH As Long
Dim zBarText As String
Dim xObj As Object

Dim xPos As Long, xtop As Long, zHasCaption As Boolean
Dim xStart As Integer, posCtr As Integer, xTotal As Long, chkSum As Long
Private Const ChkChar = 43
Sub BarEAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)
   Set xObj = zObj
   Init_Table
   zBarText = BarText
   zHasCaption = HasCaption
   xObj.Picture = Nothing
   
   If Not CheckCode13 Then Exit Sub
   
   BarH = zBarH * 10
   xtop = 10
   
   xObj.BackColor = vbWhite
   xObj.AutoRedraw = True
   xObj.ScaleMode = 3
   If HasCaption Then
      xObj.Height = (xObj.TextHeight(zBarText) + BarH + 25) * Screen.TwipsPerPixelY
   Else
      xObj.Height = (BarH + 20) * Screen.TwipsPerPixelY
   End If
   xObj.Width = (((Len(zBarText)) * 8)) * 20
   
   Paint_Bar13 zBarText
   zObj.Picture = zObj.Image
End Sub
Sub BarEAN8(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)
   Set xObj = zObj
   Init_Table
   zBarText = BarText
   zHasCaption = HasCaption
   xObj.Picture = Nothing
   
   If Not CheckCode8 Then Exit Sub
   
   BarH = zBarH * 10
   xtop = 10
   
   xObj.BackColor = vbWhite
   xObj.AutoRedraw = True
   xObj.ScaleMode = 3
   
   If HasCaption Then
      xObj.Height = (xObj.TextHeight(zBarText) + BarH + 25) * Screen.TwipsPerPixelY
   Else
      xObj.Height = (BarH + 20) * Screen.TwipsPerPixelY
   End If
   'xObj.Height = (xObj.TextHeight(zBarText) + BarH + 25) * Screen.TwipsPerPixelY
   xObj.Width = (((Len(zBarText)) * 8) + 20) * 20 'Screen.TwipsPerPixelX
  
   Paint_Bar8 zBarText
   zObj.Picture = zObj.Image
End Sub
Function CheckCode13() As Boolean
    Dim ii As Integer
    If Len(zBarText) <> 12 Then
        Err.Raise vbObjectError + 513, "EAN-13", _
          "Should be 12 Digit Numbers"
        GoTo Err_Found
    End If
    For ii = 1 To Len(zBarText)
        If InStr("0123456789", Mid(zBarText, ii, 1)) = 0 Then
            Err.Raise vbObjectError + 513, "EAN-13", _
              "An Invalid Character Found in Bar Text"
           GoTo Err_Found
        End If
    Next
    CheckCode13 = True
    Exit Function
Err_Found:
    CheckCode13 = False
End Function
Function CheckCode8() As Boolean
    Dim ii As Integer
    If Len(zBarText) <> 7 Then
        Err.Raise vbObjectError + 513, "EAN-8", _
          "Should be 7 Digit Numbers"
        GoTo Err_Found
    End If
    For ii = 1 To Len(zBarText)
        If InStr("0123456789", Mid(zBarText, ii, 1)) = 0 Then
            Err.Raise vbObjectError + 513, "EAN-8", _
              "An Invalid Character Found in Bar Text"
           GoTo Err_Found
        End If
    Next
    CheckCode8 = True
    Exit Function
Err_Found:
    CheckCode8 = False
End Function
Private Sub Paint_Bar13(ByVal xstr As String)
    Dim ii As Long, jj As Integer, ctr As Integer, xEven As Boolean, xParity As String
 
    xTotal = 0
    xPos = 5
    
    If HasCaption Then
        xObj.CurrentX = xPos
        xObj.CurrentY = 5 + BarH
        
        xObj.Print Mid(xstr, 1, 1)
    End If
    Draw_Bar "101", True
    
    xObj.CurrentY = 15 + BarH
    xParity = Parity(CInt(Mid(xstr, 1, 1)))
    
    
    For ii = 1 To Len(xstr)
        If ((Len(xstr) + 1) - ii) Mod 2 = 0 Then
           xTotal = xTotal + (CInt(Mid(xstr, ii, 1)))
        Else
           xTotal = xTotal + CInt(Mid(xstr, ii, 1) * 3)
        End If
        If ii = 8 Then
           Draw_Bar "01010", True
        End If
        jj = CInt(Mid(xstr, ii, 1))
        If ii > 1 And ii < 8 Then
           Draw_Bar CStr(IIf(Mid(xParity, ii - 1, 1) = "E", LeftHand_Even(jj), LeftHand_Odd(jj))), False
        ElseIf ii > 1 And ii >= 8 Then
           Draw_Bar CStr(Right_Hand(jj)), False
        End If
    Next
    chkSum = 0
    jj = xTotal Mod 10
    If jj <> 0 Then
       chkSum = 10 - jj
    End If
    Draw_Bar CStr(Right_Hand(chkSum)), False
    Draw_Bar "101", True
    
   If zHasCaption Then
        xObj.CurrentX = 23
        xObj.CurrentY = 10 + BarH
        xObj.Print Mid(xstr, 2, 6)
        
        xObj.CurrentX = 68
        xObj.CurrentY = 10 + BarH
        xObj.Print Mid(xstr, 8, 6) & chkSum
    End If
End Sub
Private Sub Paint_Bar8(ByVal xstr As String)
    Dim ii As Long, jj As Integer, ctr As Integer, xEven As Boolean, xParity As String
 
    xTotal = 0
    xPos = 5
    
    
    Draw_Bar "101", True
    
    xObj.CurrentX = xPos
    xObj.CurrentY = 15 + BarH
    xParity = Parity(7) 'CInt(Mid(xstr, 1, 1)))
    
    
    For ii = 1 To Len(xstr)
        If ((Len(xstr) + 1) - ii) Mod 2 = 0 Then
           xTotal = xTotal + (CInt(Mid(xstr, ii, 1)))
        Else
           xTotal = xTotal + CInt(Mid(xstr, ii, 1) * 3)
        End If
        If ii = 5 Then
           Draw_Bar "01010", True
        End If
        jj = CInt(Mid(xstr, ii, 1))
        If ii < 5 Then
           Draw_Bar CStr(LeftHand_Odd(jj)), False
        ElseIf ii >= 5 Then
           Draw_Bar CStr(Right_Hand(jj)), False
        End If
    Next
    chkSum = 0
    jj = xTotal Mod 10
    If jj <> 0 Then
       chkSum = 10 - jj
    End If
    Draw_Bar CStr(Right_Hand(chkSum)), False
    Draw_Bar "101", True
    
    If zHasCaption Then
        xObj.CurrentX = 23
        xObj.CurrentY = 10 + BarH
        xObj.Print Mid(xstr, 1, 4)
        
        xObj.CurrentX = 53
        xObj.CurrentY = 10 + BarH
        xObj.Print Mid(xstr, 5, 4) & chkSum
    End If
End Sub

Private Sub Draw_Bar(Encoding As String, Guard As Boolean)
    Dim ii As Integer
    For ii = 1 To Len(Encoding)
        xPos = xPos + 1
        xObj.Line (xPos + 10, xtop)-(xPos + 10, xtop + BarH + IIf(Guard, 5, 0)), IIf(Mid(Encoding, ii, 1), vbBlack, vbWhite)
    Next
End Sub
Private Sub Init_Table()
    LeftHand_Odd = Array("0001101", "0011001", "0010011", "0111101", "0100011", "0110001", "0101111", "0111011", "0110111", "0001011")
    LeftHand_Even = Array("0100111", "0110011", "0011011", "0100001", "0011101", "0111001", "0000101", "0010001", "0001001", "0010111")
    Right_Hand = Array("1110010", "1100110", "1101100", "1000010", "1011100", "1001110", "1010000", "1000100", "1001000", "1110100")
    Parity = Array("OOOOOO", "OOEOEE", "OOEEOE", "OOEEEO", "OEOOEE", "OEEOOE", "OEEEOO", "OEOEOE", "OEOEEO", "OEEOEO")
End Sub
