VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRowGuid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'RETURNS:GUID if successful; blank string otherwise.
'Unlike the GUIDS in the reGIStry, this function returns GUID
'without "-" characters.See comments for how to modify if you
'want the dash.
Private Type Guid
    PartOne As Long
    PartTwo As Long
    PartThree As Long
    PartFour(8) As Byte
End Type

Private Declare Function CoCreateGuid _
                Lib "ole32.dll" (pguid As Guid) As Long
Private Declare Function StringFromGUID2 _
                Lib "ole32.dll" (rguid As Any, _
                                 ByVal lpstrClsId As Long, _
                                 ByVal cbMax As Long) As Long


Public Function Guid() As String
    Dim uGUID As Guid
    Dim sGUID As String
    Dim bGUID() As Byte
    Dim lLen As Long
    Dim RetVal As Long
    lLen = 40
    bGUID = String$(lLen, 0)
    CoCreateGuid uGUID '把结构转换为一个可显示的字符串
    RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
    sGUID = bGUID

    If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then RetVal = RetVal - 1
    Guid = Left$(sGUID, RetVal)

End Function

