Attribute VB_Name = "INI"
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'MsgBox GetINI("User Info", "Username")'
'SetINI "User Info", "Username", "hier je username"'

Function GetINI(strMain As String, strSub As String) As String
    Dim strBuffer As String
    Dim lngLen As Long
    Dim lngRet As Long
    strBuffer = Space(100)
    lngLen = Len(strBuffer)
    lngRet = GetPrivateProfileString(strMain, strSub, vbNullString, strBuffer, lngLen, App.Path & "/users.ini")
    GetINI = Left(strBuffer, lngRet)
End Function

Sub SetINI(strMain As String, strSub As String, strValue As String)
    WritePrivateProfileString strMain, strSub, strValue, App.Path & "/users.ini"
End Sub

        

