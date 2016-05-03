VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUpLoadData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function PostExpressData(ByVal Site As String, ByVal ExpressNO As String, ByVal strState As String) As String

    Dim strJSON As String
    
    strJSON = "{"
    strJSON = strJSON & """Type"":""Save_ExpressDetail"""
    strJSON = strJSON & ",""SiteCode"":""" & Site & """"
    strJSON = strJSON & ",""ExpressNO"":""" & ExpressNO & """"
    strJSON = strJSON & ",""State"":""" & strState & """"
    strJSON = strJSON & "}"
    Dim strResult As String
    strResult = modPostData_Core.PostData("/plugin/venderexpress.asp", strJSON)
    Debug.Print strResult
    Dim dicResult As Scripting.Dictionary
    Set dicResult = JSON.Parse(strResult)
    
    If dicResult.Exists("OK") Then
    
        
    
    Else
    
    End If



End Function
