VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransformPWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function TransFormPWD(ByVal strPWD As String) As String
    Dim i As Integer
    i = 0
    Dim sb As String

    For i = 0 To Len(strPWD) - 1

        Select Case i Mod 6

            Case 0:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 8)

            Case 1:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 3)
     
            Case 2:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 9)

            Case 3:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 6)
         
            Case 4:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 2)

            Case 5:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 3)
                      
        End Select

    Next

    TransFormPWD = sb
End Function

Public Function deTransFormPWD(ByVal strPWD As String) As String
    Dim i As Integer
    i = 0
    Dim sb As String

    For i = 0 To Len(strPWD) - 1

        Select Case i Mod 6

            Case 0:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) + 8)

            Case 1:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) + 3)
     
            Case 2:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) + 9)

            Case 3:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) + 6)
         
            Case 4:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) + 2)

            Case 5:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) + 3)
                      
        End Select

    Next

    deTransFormPWD = sb
End Function
