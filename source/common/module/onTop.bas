Attribute VB_Name = "modOnTop"
Option Explicit

#If Win16 Then
    Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#Else
    Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If


'*************************************************************************
'* Function: KeepOnTop(F As Form)
'*
'*
'*************************************************************************
'* Description: Keep form on top.
'*
'*
'*************************************************************************
'* Parameters: Form Control
'*
'*************************************************************************
'* Notes: The SetWindowPos API call gets turned off if the form is
'*        minimized.  Put this code in the resize event to make sure
'*        your form stays on top.
'*
'*************************************************************************
'* Returns:
'*************************************************************************
Sub KeepOnTop(F As Form)
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

    SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

