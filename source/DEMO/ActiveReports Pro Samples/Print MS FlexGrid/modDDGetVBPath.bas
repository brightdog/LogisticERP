Attribute VB_Name = "modDDGetVBPath"
Option Explicit

Private Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpszValueName$, ByVal lpdwRes&, lpdwType&, ByVal lpDataBuff$, nSize&)
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Public Function GetVBPath() As String
Dim s As String * 255, sVBEntry As String
Dim lVBKey As Long, lType As Long, lLen As Long, lRC As Long
Dim sPath As String
Const HKEY_LOCAL_MACHINE = &H80000002
Const KEY_QUERY_VALUE = &H1&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
    
    ' Search for VB5 path
    sVBEntry = "SOFTWARE\Microsoft\Visual Basic\5.0"
    sPath = ""
    lLen = Len(s)
    lRC = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sVBEntry, 0, KEY_READ, lVBKey)
    If lRC = 0 Then
        ' RegOpenKeyEx succeeded
        lRC = RegQueryValueEx(lVBKey, "InstallDir", 0, lType, s, lLen)
        If lRC = 0 Then sPath = Left$(s, lLen - 1)  'RegQueryValueEx succeded
        RegCloseKey lVBKey
    End If

    If sPath <> "" Then
        GetVBPath = sPath
        Exit Function
    End If
    
    ' vb5 key not found, search for vb6 path
    sVBEntry = "Software\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic"
    lLen = Len(s)
    lRC = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sVBEntry, 0, KEY_READ, lVBKey)
    If lRC = 0 Then
        lRC = RegQueryValueEx(lVBKey, "ProductDir", 0, lType, s, lLen)
        If lRC = 0 Then
            sPath = Left$(s, lLen - 1)  ' vb6 key found
        Else
            sPath = InputBox("Please enter a valid path to NWind.mdb", "VB Path not found", "E:\Microsoft Visual Studio\VB98")
        End If
        RegCloseKey lVBKey
    End If
    GetVBPath = sPath
End Function
