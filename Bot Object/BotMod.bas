Attribute VB_Name = "BotMod"
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long


Public Function CountTokens(ByVal val As String, ByVal tok As String) As Integer
    Dim count As Integer
    Dim x As Integer
    count = 0
    x = InStr(val, tok)
    Do Until x = 0
        val = Mid(val, x + 1)
        count = count + 1
        x = InStr(val, tok)
    Loop
    CountTokens = count + 1
End Function
Public Function getNextToken(ByRef szTemp$, ByVal szTok$)
    Dim Y As Integer
    Y = InStr(szTemp$, szTok$)
    If Y <> 0 Then
        getNextToken = Left(szTemp$, Y - 1)
        szTemp$ = Mid(szTemp$, Y + 1)
    Else
        getNextToken = szTemp$
    End If
End Function
