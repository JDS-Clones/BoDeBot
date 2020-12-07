Attribute VB_Name = "BoDeII"
Option Explicit
''***Functions
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long

Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (retval As Byte, ByVal Ptr As Long) As Long
Declare Function StrToPtr Lib "kernel32" Alias "lstrcpyW" (ByVal Ptr As Long, Source As Byte) As Long
Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" (retval As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long
Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Declare Function PtrToArray Lib "kernel32" Alias "lstrcpyW" (retval As Byte, ByVal Ptr As Long) As Long
Declare Function GetStrPtr Lib "kernel32" Alias "lstrcpyW" (Dest As Byte, Source As Byte) As Long

Public WINMIN As Boolean 'use this to avoid the taskbar infinite loop
Public WINDOW_STATE As Integer
''***Constants
Public Const EM_SCROLLCARET = &HB7

Public Sub UpdateRTB(thebox As Control, thetext As String, Optional thecolor As Integer)
    Dim XX As Integer
    thebox.SelStart = Len(thebox)
    thebox.SelColor = QBColor(thecolor)
    thebox.SelBold = MDI.NetClient.FontBold
    thebox.SelFontName = MDI.NetClient.FontName
    thebox.SelFontSize = MDI.NetClient.FontSize
    thebox.SelText = thetext & vbCrLf
    XX = SendMessage(thebox.hWnd, EM_SCROLLCARET, ByVal 0, ByVal 0)
End Sub

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

Public Function DecodeLongIPAddr(ByVal LongIPAddr As String) As String
    Dim HiWord!, LoWord!, Nibble_1, Nibble_2, Nibble_3, Nibble_4, IPAddr As String
    HiWord! = Int(LongIPAddr / 65536)
    LoWord! = LongIPAddr - HiWord! * 65536
    Nibble_1 = Int(HiWord! / 256)
    Nibble_2 = HiWord! - Nibble_1 * 256
    Nibble_3 = Int(LoWord! / 256)
    Nibble_4 = LoWord! - Nibble_3 * 256
    IPAddr = Trim(Str(Nibble_1)) & "."
    IPAddr = IPAddr & Trim(Str(Nibble_2)) & "."
    IPAddr = IPAddr & Trim(Str(Nibble_3)) & "."
    IPAddr = IPAddr & Trim(Str(Nibble_4))
    DecodeLongIPAddr = IPAddr
End Function
Public Function EncodeIPAddr(IPAddr As String) As String
    Dim DStart, EIP, DStop, ThisByte, HiWord!, LoWord!, LongIP
    DStart = 1
    EIP = ""
    Do
       DStop = InStr(DStart, IPAddr & ".", ".")
       ThisByte = Hex(val(Mid$(IPAddr & ".", DStart, DStop - DStart)))
       EIP = EIP & IIf(Len(ThisByte) = 1, "0" & ThisByte, ThisByte)
       DStart = DStop + 1
    Loop Until DStart >= Len(IPAddr & ".")
    HiWord! = val("&H" & Mid(EIP, 1, 2)) * 256! + val("&H" & Mid(EIP, 3, 2))
    LoWord! = val("&H" & Mid(EIP, 5, 2)) * 256! + val("&H" & Mid(EIP, 7, 2))
    LongIP = HiWord! * 65536 + LoWord!
    EncodeIPAddr = Trim$(Str$(LongIP))
End Function

Public Function existsInList(ml As ListView, szTok As String) As Boolean
    existsInList = False 'init to false
    Dim num As Integer
        For num = 1 To ml.ListItems.count
            If LCase$(ml.ListItems(num).key) = LCase$(szTok) Then
                existsInList = True
            End If
        Next num
End Function



Public Sub startDCCSend(ByVal szNick As String)
    Dim szFile As String, szRest As String, param1$
    Dim sfile As DCCSend
    MDI.CMD.Filter = "All Files (*.*)|*.*"
    MDI.CMD.DialogTitle = "DCC Send"
    MDI.CMD.flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNLongNames
    MDI.CMD.ShowOpen
    szRest = MDI.CMD.filename
    Do Until szFile = szRest
        Set sfile = New DCCSend
        Load sfile
        param1$ = CStr(sfile.hWnd)
        MDI.NetClient.DccWins.Add sfile, param1$
        sfile.Show
        szFile = getNextToken(szRest, Chr$(0))
        MDI.NetClient.DCCSendFile szNick, szFile, sfile
    Loop
    
    

End Sub
Public Sub startDCCChat(ByVal szNick As String)
    Dim schat As DCCChat
    Dim param1$
    Set schat = New DCCChat
    Load schat
    param1$ = CStr(schat.hWnd)
    MDI.NetClient.DccWins.Add schat, param1$
    schat.Show
    MDI.NetClient.DCCSendchat szNick, schat
End Sub

Public Function ReturnFileName(ByVal vData As String) As String
Dim iLastSlash  As Integer
Dim i           As Integer
 
    'Assuming all goes well, anything after the
    'last backslash is part of the filename and anything
    'before the last backslash is the folder
 
    iLastSlash = 0
    Do
        If Len(vData) > iLastSlash + 1 Then
            i = InStr(iLastSlash + 1, vData, "\")
        Else
            i = 0
        End If
        If (i > 0) Then iLastSlash = i
 
    Loop While (i > 0)
    ReturnFileName = Right$(RTrim$(vData), Len(RTrim$(vData)) - iLastSlash)
 
End Function
Public Function GetFromIni(ByVal szSection As String, ByVal szField As String) As String
    Dim di As Long
    Dim retval As String * 4096
    ' added by sk8
    Dim ret As Long
    Dim retstr As String
    Dim file As String
    retstr = Space$(255)
    file = App.Path & "\bodebot2.ini"
    ret = GetPrivateProfileString(UCase$(szSection), UCase$(szField), "", retstr, Len(retstr), file)
    GetFromIni = Left(retstr, ret)
End Function
