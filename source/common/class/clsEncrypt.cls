VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Encode(ByVal s As String) As String
    On Error GoTo Err

    If Len(s) = 0 Then Exit Function
    Dim Buff()     As Byte
    Buff = StrConv(s, vbFromUnicode)
    Dim i     As Long
    Dim j     As Byte
    Dim k     As Byte, m       As Byte
    Dim mstr     As String
    mstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyz"
    Dim outs     As String
    i = UBound(Buff) + 1
    outs = Space$(2 * i)
    Dim Temps     As String

    For i = 0 To UBound(Buff)
        Randomize Time
        j = CByte(5 * (Math.Rnd()) + 0)
        Buff(i) = Buff(i) Xor j
        k = Buff(i) Mod Len(mstr)
        m = Buff(i) \ Len(mstr)
        m = m * 2 ^ 3 + j
        Temps = Mid$(mstr, k + 1, 1) + Mid$(mstr, m + 1, 1)
        Mid$(outs, 2 * i + 1, 2) = Temps
    Next

    Encode = outs
    Exit Function
Err:
    Encode = "ERR!"
End Function
   
'½âÃÜº¯Êý
Public Function Decode(ByVal s As String) As String
    On Error GoTo Err
    Dim i     As Long
    Dim j     As Byte
    Dim k     As Byte
    Dim m     As Byte
    Dim mstr     As String
    mstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyz"
    Dim t1     As String, t2       As String
    Dim Buff()     As Byte
    Dim n     As Long
    n = 0

    For i = 1 To Len(s) Step 2
        t1 = Mid$(s, i, 1)
        t2 = Mid$(s, i + 1, 1)
        k = InStr(1, mstr, t1) - 1
        m = InStr(1, mstr, t2) - 1
        j = m \ 2 ^ 3
        m = m - j * 2 ^ 3
        ReDim Preserve Buff(n)
        Buff(n) = j * Len(mstr) + k
        Buff(n) = Buff(n) Xor m
        n = n + 1
    Next

    Decode = StrConv(Buff, vbUnicode)
    Exit Function
Err:
    Decode = "ERR!"
End Function
