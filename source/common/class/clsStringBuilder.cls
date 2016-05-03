VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Binary

Private Declare Sub RtlMoveMemory _
                Lib "kernel32" (dst As Any, _
                                src As Any, _
                                ByVal nBytes&)
Private Declare Function SysAllocStringByteLen& _
                Lib "oleaut32" (ByVal olestr&, _
                                ByVal BLen&)

Private plngStringLen As Long
Private plngBufferLen As Long
Private pstrBuffer As String

Public Sub Append(Text As String)
    Dim lngText As Long
    Dim strTemp As String
    Dim lngVPointr As Long

    lngText = Len(Text)

    If lngText > 0 Then
        If (plngStringLen + lngText) > plngBufferLen Then
            plngBufferLen = (plngStringLen + lngText) * 2&
            strTemp = AllocString04(plngBufferLen)

            '***copymemory might be faster than this
            Mid$(strTemp, 1&) = pstrBuffer

            '***Alternate pstrBuffer = strTemp
            '***switch pointers instead of slow =
            lngVPointr = StrPtr(pstrBuffer)
            RtlMoveMemory ByVal VarPtr(pstrBuffer), ByVal VarPtr(strTemp), 4&
            RtlMoveMemory ByVal VarPtr(strTemp), lngVPointr, 4&

            'Debug.Print "plngBufferLen: " & plngBufferLen
        End If

        Mid$(pstrBuffer, plngStringLen + 1&) = Text
        plngStringLen = plngStringLen + lngText
    End If

End Sub

Public Function ToString() As String
    ToString = Left$(pstrBuffer, plngStringLen)
End Function

Public Property Let value(ByRef newValue As String)
    Call Clear
    Me.Append newValue

End Property

Private Function AllocString04(ByVal lSize As Long) As String
    ' http://www.xbeat.net/vbspeed/
    ' by Jory, jory@joryanick.com, 20011023
    RtlMoveMemory ByVal VarPtr(AllocString04), SysAllocStringByteLen(0&, lSize + lSize), 4&
End Function

Private Sub Clear()
    '***do not clear the buffer to save allocation time
    '***if you use the function multiple times
    plngStringLen = 0&

    plngBufferLen = 0& 'clear the buffer
    pstrBuffer = vbNullString 'clear the buffer
End Sub

Public Function Format(ByVal FormatStr As String, _
                       ParamArray args() As Variant) As clsStringBuilder
    Dim pos As Long
    Dim length As Long
    Dim ch As String
    Dim p As Long
    Dim i As Long
    Dim j As Long
    Dim s As String
    Dim curStr As String
    Dim pad As Long
    Dim index As Long
    Dim leftjustify As Boolean
    Dim width As Long
    Dim fmt As String
    Dim col As New VBA.Collection

    For i = LBound(args) To UBound(args)

        If Right$(TypeName(args(i)), 1) = ")" Then

            For j = LBound(args(i)) To UBound(args(i))
                col.Add args(i)(j)
            Next

        Else
            col.Add args(i)
        End If

    Next

    s = FormatStr

    length = Len(s)
    pos = 0
    ch = Chr$(0)

    Me.value = ""

    Do While (True)
        p = pos
        i = pos

        Do While (pos < length)
            ch = Mid$(s, pos + 1, 1)
            pos = pos + 1

            If (ch = "}") Then
                If (pos < length And Mid$(s, pos + 1, 1) = "}") Then ' treat asescape character fro }}
                    pos = pos + 1
                Else

                    formatError "} unexcept in pos:" & pos
                End If
            End If

            If (ch = "{") Then
                If (pos < length And Mid$(s, pos + 1, 1) = "{") Then ' treat as escape character for {{
                    pos = pos + 1
                Else
                    pos = pos - 1
                    Exit Do
                End If
            End If

            Mid$(s, i + 1, 1) = ch
            i = i + 1
        Loop

        If (i > p) Then AppendRepeat Mid$(s, p + 1, i - p)
        If (pos = length) Then Exit Do

        pos = pos + 1
        ch = Mid$(s, pos + 1, 1)

        If (pos = length Or ch < "0" Or ch > "9") Then formatError "wrong number in pos:" & pos

        index = 0

        Do
            index = index * 10 + (Asc(ch) - Asc("0"))
            pos = pos + 1

            If (pos = length) Then formatError "unfinished"
            ch = Mid$(s, pos + 1, 1)
        Loop While (ch >= "0" And ch <= "9" And index < 1000000)

        If (index > col.Count - 1) Then formatError "index of args out of range"

        ch = Mid$(s, pos + 1, 1)

        Do While (pos < length And ch = " ")
            pos = pos + 1
            ch = Mid$(s, pos + 1, 1)
        Loop

        leftjustify = False
        width = 0

        If (ch = ",") Then
            pos = pos + 1

            Do While (pos < length And Mid$(s, pos + 1, 1) = " ")
                pos = pos + 1
            Loop

            If (pos = length) Then formatError "unfinished"
            ch = Mid$(s, pos + 1, 1)

            If (ch = "-") Then
                leftjustify = True
                pos = pos + 1

                If (pos = length) Then formatError "unfinished"
                ch = Mid$(s, pos + 1, 1)
            End If

            If (ch < "0" Or ch > "9") Then formatError "wrong number in pos:" & pos

            Do
                width = width * 10 + (Asc(ch) - Asc("0"))
                pos = pos + 1

                If (pos = length) Then formatError "unfinished"
                ch = Mid$(s, pos + 1, 1)
            Loop While (ch >= "0" And ch <= "9" And width < 1000000)

            ch = Mid$(s, pos + 1, 1)

            Do While (pos < length And ch = " ")
                pos = pos + 1
                ch = Mid$(s, pos + 1, 1)
            Loop

        End If

        If ch = ":" Then
            pos = pos + 1
            p = pos
            i = pos

            Do While True

                If (pos = length) Then formatError "unfinished"
                ch = Mid$(s, pos + 1, 1)
                pos = pos + 1

                If (ch = "{") Then

                    If (pos < length And Mid$(s, pos + 1, 1) = "}") Then ' treat asescape character fro }}
                        pos = pos + 1
                    Else

                        formatError "{ unexcept in pos:" & pos
                    End If

                End If

                If (ch = "}") Then
                    If (pos < length And Mid$(s, pos + 1, 1) = "{") Then ' treat as escape character for {{
                        pos = pos + 1
                    Else
                        pos = pos - 1
                        Exit Do
                    End If
                End If

                Mid$(s, i, 1) = ch
                i = i + 1

            Loop

            If (i > p) Then fmt = Mid$(s, p, i - p)
        End If

        If (ch <> "}") Then formatError "} except in pos:" & pos

        If fmt = "" Then
            curStr = CStr(col(index + 1))
        Else
            curStr = VBA.Format$(col(index + 1), fmt)
        End If

        pos = pos + 1

        pad = width - Len(curStr)

        If (Not leftjustify And pad > 0) Then Call AppendRepeat(" ", pad)
        AppendRepeat curStr

        If (leftjustify And pad > 0) Then Call AppendRepeat(" ", pad)

    Loop

    Set Format = Me
End Function

Private Sub formatError(ByRef msg As String)
    Err.Raise vbObjectError, , msg
End Sub

Private Sub AppendRepeat(s As String, _
                         Optional cnt As Long = 1)
    Dim i As Long

    For i = 1 To cnt
        Me.Append s
    Next

End Sub

