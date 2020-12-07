VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBotFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************************************
'   IRC Bot Component
'       Author: Nicholas J. Felmlee
'       Email: Nick@Felmlee.Com
'       Last Revision: 07/18/98
'
'       Internal Dependencies:
'           WskSock.bas
'           Crypt.bas
'           CBotFactory.cls
'           CBotConn.cls
'
'       External Dependencies:
'           -none-
'
'**************************************************************

Option Explicit

'local variable(s) to hold property value(s)
Private mvarDisableDefaultOps As Boolean 'local copy
Private mvarMaxConnections As Integer 'local copy
Private mvarlpIRC_SOCKET As Long 'local copy
Private mvarhWnd As Long 'local copy
Private mvarWndProc As Long

Public Event onUserJoin(ByVal szUser As String, ByVal szHost As String, ByVal dwCount As Integer)
Public Event onUserPart(ByVal szName As String, ByVal dwCount As Integer)
Public Event onUserCommand(ByVal UserOb As Object, ByVal szCommand As String, ByVal szParams As String)


Private BotConns As New Collection

'internal levels
Private PRIV_OP As Integer
Private PRIV_DEOP As Integer
Private PRIV_PV As Integer
Private PRIV_MV As Integer
Private PRIV_KICK As Integer
Private PRIV_BAN As Integer
Private PRIV_UNBAN As Integer
Private PRIV_SAY As Integer
Private PRIV_KICKBAN As Integer
Private PRIV_PERMBAN As Integer
Private PRIV_TOPIC As Integer
Private PRIV_JOIN As Integer
Private PRIV_PART As Integer
Private PRIV_ADDUSER As Integer
Private PRIV_DELUSER As Integer
Private PRIV_LSUSERS As Integer
Private PRIV_SETPERMS As Integer

Private Function GetChannel(ByVal szLine As String) As String
    'channel is always the second argument in a user command
    GetChannel = getNextToken(szLine, " ")
End Function

Private Function GetPrivProfString(ByVal szSection As String, ByVal szKey As String, ByVal szFile As String)
    Dim ret As Long
    Dim retstr As String
    Dim file As String
    retstr = Space$(255)
    file = App.Path & "\" & szFile
    ret = GetPrivateProfileString(UCase$(szSection), UCase$(szKey), "", retstr, Len(retstr), file)
    GetPrivProfString = Left(retstr, ret)
End Function
Private Sub WritePrivProfString(ByVal szString As String, ByVal szSection As String, ByVal szKey As String, ByVal szFile As String)
    Dim ret As Long
    'Dim retstr As String
    Dim file As String
    file = App.Path & "\" & szFile
    ret = WritePrivateProfileString(UCase$(szSection), UCase$(szKey), szString, file)
End Sub

Private Function GetPrivProfSection(ByVal szSection As String)
    Dim ret As Long
    Dim retstr As String
    Dim file As String
    retstr = Space$(255)
    file = App.Path & "\pass.txt"
    ret = GetPrivateProfileSection(szSection, retstr, Len(retstr), file)
    'ret = GetPrivateProfileString(UCase$(szSection), UCase$(szKey), "", retstr, Len(retstr), file)
    GetPrivProfSection = Left(retstr, ret)
End Function
Public Property Let hWnd(ByVal vData As Long)
    mvarhWnd = vData
End Property


Public Property Get hWnd() As Long
    hWnd = mvarhWnd
End Property
Public Property Let WndProc(ByVal vData As Long)
    mvarWndProc = vData
End Property
Public Property Get WndProc() As Long
    WndProc = mvarWndProc
End Property


Public Sub GetNewConnection(szNick As String)
    Dim MySock As Long
    Dim addrlen&
    Dim ret&
    Dim sa As sockaddr
    Dim dwPort As Integer
    addrlen = sockaddr_size
    ret = getsockname(mvarlpIRC_SOCKET, sa, addrlen)
    dwPort = ListenToFreePort(MySock)
    Call SendData(mvarlpIRC_SOCKET, "PRIVMSG " & szNick & " :" & Chr(1) & "DCC CHAT chat " & htonl(sa.sin_addr) & " " & dwPort & Chr(1) & vbCrLf)
End Sub


Public Property Let lpIRC_SOCKET(ByVal vData As Long)
    mvarlpIRC_SOCKET = vData
End Property


Public Property Get lpIRC_SOCKET() As Long
    lpIRC_SOCKET = mvarlpIRC_SOCKET
End Property


'
Private Function ListenToFreePort(ByRef MySock As Long) As Integer
    Dim sockin As sockaddr
    Dim addrlen&, s&, Dummy&, SelectOps&
    sockin = saZero     'zero out the structure
    sockin.sin_family = AF_INET
    
    'sockin.sin_port = htons(0)
    sockin.sin_port = 0
    If sockin.sin_port = INVALID_SOCKET Then
        ListenToFreePort = INVALID_SOCKET
        Exit Function
    End If
    
    sockin.sin_addr = htonl(INADDR_ANY)
    If sockin.sin_addr = INADDR_NONE Then
        ListenToFreePort = INVALID_SOCKET
        Exit Function
    End If
    
    s = socket(PF_INET, SOCK_STREAM, 0)
    If s < 0 Then
        ListenToFreePort = INVALID_SOCKET
        Exit Function
    End If
    
    If bind(s, sockin, sockaddr_size) Then
        If s > 0 Then
            Dummy = closesocket(s)
        End If
        ListenToFreePort = INVALID_SOCKET
        Exit Function
    End If
    
    SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    If WSAAsyncSelect(s, mvarhWnd, ByVal 1025, ByVal SelectOps) Then
        If s > 0 Then
            Dummy = closesocket(s)
        End If
        ListenToFreePort = SOCKET_ERROR
        Exit Function
    End If
    
    If listen(s, 1) Then
        If s > 0 Then
            Dummy = closesocket(s)
        End If
        ListenToFreePort = INVALID_SOCKET
        Exit Function
    End If
    Dim sadd As sockaddr
    MySock = s
    addrlen = sockaddr_size
    Call getsockname(s, sadd, addrlen)
    ListenToFreePort = htons(sadd.sin_port)
End Function




Public Function ListUsers() As Object
End Function

Public Sub DelUser(ByVal szUser As String)
End Sub

Public Sub AddUser(ByVal szUser As String, ByVal szPass As String, ByVal dwLevel As Integer)
End Sub

Public Sub KickUser(ByVal szUser As String)
End Sub

Public Property Let MaxConnections(ByVal vData As Integer)
    mvarMaxConnections = vData
End Property


Public Property Get MaxConnections() As Integer
    MaxConnections = mvarMaxConnections
End Property



Public Property Let DisableDefaultOps(ByVal vData As Boolean)
    mvarDisableDefaultOps = vData
End Property


Public Property Get DisableDefaultOps() As Boolean
    DisableDefaultOps = mvarDisableDefaultOps
End Property



Public Function Hook(ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    Dim x As Long, a As String, i As Long
    Dim ReadBuffer(1) As Byte
    Dim SendD As String
    Dim newsock As Long
    Dim saddr As sockaddr
    Static CurrentLine As String
    Static was_delim As Boolean
    Select Case msg
        Case 1025
            '#We have connected
            If lp = FD_CONNECT Then
                Debug.Print "Bot Connect"
        
            '#Data recieved yippie
            ElseIf lp = FD_READ Then
                x = recv(wp, ReadBuffer(0), 1, 0)
                If x > 0 Then
                    a = StrConv(ReadBuffer, vbUnicode)
                    a = CStr(Chr$(Asc(a)))
                    If Not ((a = Chr$(10)) Or (a = Chr$(13)) Or (a = (Chr$(10) & Chr$(13)))) Then
                        CurrentLine = CurrentLine & a
                        was_delim = False
                    Else
                        If Not was_delim Then
                            ProcessBotInput wp, CurrentLine
                            CurrentLine = ""
                            was_delim = True
                        End If
                    End If
                End If
            
            ElseIf lp = FD_CLOSE Then
                Debug.Print "Bot Close"
                x = closesocket(wp)
                RemoveConn wp
            
            ElseIf lp = FD_ACCEPT Then
                Debug.Print "Bot Accept"
                newsock = accept(wp, saddr, sockaddr_size)
                Call closesocket(wp)
                Call SendData(newsock, "############################" & vbCrLf)
                Call SendData(newsock, "# BoDeBoT (c) 1998         #" & vbCrLf)
                Call SendData(newsock, "# PoWeReD by VB5           #" & vbCrLf)
                Call SendData(newsock, "############################" & vbCrLf)
                Call SendData(newsock, Now & vbCrLf)
                Call SendData(newsock, "Login:" & vbCrLf)
                AddConn newsock
            End If
            Hook = 0
        Case Else
            Hook = mvarWndProc
    End Select

End Function
Private Sub RemoveConn(ByVal ConnID As Long)
    Dim szKey As String
    szKey = CStr(ConnID)
    'Set BotConns(szKey) = Nothing
    If BotConns.Item(szKey).Status >= 2 Then
        SendToPartyLine "** " & BotConns.Item(szKey).User & " left the party line."
        RaiseEvent onUserPart(BotConns.Item(szKey).User, BotConns.count - 1)
    End If
    BotConns.Remove szKey
End Sub

Private Sub AddConn(ByVal ConnID As Long)
    Dim ucon As CBotConn
    Dim szKey As String
    szKey = CStr(ConnID)
    Set ucon = New CBotConn
    ucon.User = ""
    ucon.Status = 0 'zero means waiting for user
    ucon.Level = 0 'zero until we get a login
    ucon.LastCommand = "Login: Waiting for Username"
    ucon.ConnID = ConnID
    BotConns.Add ucon, szKey
    'We won't fire event until we get an official login from Authenticate
    'MDI.StatusBar1.Panels(2).Text = BotConns.count & " users logged into bot."
End Sub

Private Sub ProcessBotInput(ByVal ConnID As Long, ByVal szLine As String)
    Dim baccess As Integer
    Dim sw As Integer
    Dim szKey As String
    szKey = CStr(ConnID)
    sw = BotConns.Item(szKey).Status
    Debug.Print sw
    Select Case sw
        Case 0 'was waiting for a username
            BotConns.Item(szKey).User = szLine
            BotConns.Item(szKey).Status = 1
            Call SendData(ConnID, "Pass:" & vbCrLf)
        Case 1 'was waiting for a password
            'check password file
            baccess = Authenticate(BotConns.Item(szKey).User, szLine)
            If baccess > 0 Then
                RaiseEvent onUserJoin(BotConns.Item(szKey).User, "", BotConns.count)
                BotConns.Item(szKey).Status = 2
                BotConns.Item(szKey).SetPerms GetPrivProfSection(BotConns.Item(szKey).User) 'load the perms
                Call SendData(ConnID, "Welcome " & BotConns.Item(szKey).User & vbCrLf)
                'send join message to party line
                SendToPartyLine "** " & BotConns.Item(szKey).User & " joined the party line."
            Else
                BotConns.Item(szKey).Status = 0
                Call SendData(ConnID, "Access Denied!" & vbCrLf)
                Call closesocket(ConnID)
                RemoveConn ConnID
            End If
        Case 2 'authenticated
            'Everything is cool, process commands
            ProcessUserCommands ConnID, BotConns.Item(szKey), szLine
    End Select
End Sub

Private Function Authenticate(ByVal szUser As String, ByVal szPass As String) As Integer
    '*Returns Levels of User if auhenticated
    '*else returns 0 for no access
    '*note that level 0 has no access
    Authenticate = 0
    If Mid(szUser, 1, 1) = "#" Then Exit Function
    If Decrypt(GetPrivProfString(szUser, "pass", "pass.txt")) = szPass Then
        Authenticate = 1
    End If
End Function

Private Sub ProcessUserCommands(ByVal ConnID As Long, bot As CBotConn, ByVal szLine As String)
    Dim szCommand As String, temp As String, param1$, param2$, param3$, param4$, test$
    Dim num As Integer
    Dim szKey As String
    szKey = CStr(ConnID)
    szCommand = getNextToken(szLine, " ")
    RaiseEvent onUserCommand(BotConns.Item(szKey), szCommand, szLine) 'send to event for outside script processing
    If mvarDisableDefaultOps = False Then 'speeds things up by accepting default bot ops
        Select Case LCase$(szCommand)
            Case ".op"
                BotConns.Item(szKey).LastCommand = szCommand
                If BotConns.Item(szKey).CheckChanPerms(GetChannel(szLine), PRIV_OP) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") >= 2 Then
                    param1$ = getNextToken(szLine, " ") '#channel
                    Call SendData(mvarlpIRC_SOCKET, "MODE " & param1$ & " +o " & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.op #channel nick>"
                End If
                
            Case ".deop"
                BotConns.Item(szKey).LastCommand = szCommand
                If BotConns.Item(szKey).CheckChanPerms(GetChannel(szLine), PRIV_DEOP) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") >= 2 Then
                    param1$ = getNextToken(szLine, " ") '#channel
                    Call SendData(mvarlpIRC_SOCKET, "MODE " & param1$ & " -o " & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.deop #channel nick>"
                End If
                
            Case ".+v"
                BotConns.Item(szKey).LastCommand = szCommand
                If BotConns.Item(szKey).CheckChanPerms(GetChannel(szLine), PRIV_PV) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") >= 2 Then
                    param1$ = getNextToken(szLine, " ") '#channel
                    Call SendData(mvarlpIRC_SOCKET, "MODE " & param1$ & " +v " & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.+v #channel nick>"
                End If
                
            Case ".-v"
                BotConns.Item(szKey).LastCommand = szCommand
                If BotConns.Item(szKey).CheckChanPerms(GetChannel(szLine), PRIV_MV) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") >= 2 Then
                    param1$ = getNextToken(szLine, " ") '#channel
                    Call SendData(mvarlpIRC_SOCKET, "MODE " & param1$ & " -v " & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.-v #channel nick>"
                End If
                
            Case ".+ban"
                BotConns.Item(szKey).LastCommand = szCommand
                If BotConns.Item(szKey).CheckChanPerms(GetChannel(szLine), PRIV_BAN) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") >= 2 Then
                    param1$ = getNextToken(szLine, " ") '#channel
                    Call SendData(mvarlpIRC_SOCKET, "MODE " & param1$ & " +b " & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.+ban #channel params>"
                End If
                
            Case ".-ban"
                BotConns.Item(szKey).LastCommand = szCommand
                If BotConns.Item(szKey).CheckChanPerms(GetChannel(szLine), PRIV_UNBAN) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") >= 2 Then
                    param1$ = getNextToken(szLine, " ") '#channel
                    Call SendData(mvarlpIRC_SOCKET, "MODE " & param1$ & " -b " & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.-ban #channel params>"
                End If
            
            Case ".banlist"
                BotConns.Item(szKey).LastCommand = szCommand
            
            Case ".kick"
                BotConns.Item(szKey).LastCommand = szCommand
                If BotConns.Item(szKey).CheckChanPerms(GetChannel(szLine), PRIV_KICK) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") >= 2 Then
                    param1$ = getNextToken(szLine, " ") '#channel
                    param2$ = getNextToken(szLine, " ") '#channel
                    Call SendData(mvarlpIRC_SOCKET, "KICK " & param1$ & " " & param2$ & " :" & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.kick #channel nickname kickmessage>"
                End If
                
            Case ".kickban"
                BotConns.Item(szKey).LastCommand = szCommand
                If BotConns.Item(szKey).CheckChanPerms(GetChannel(szLine), PRIV_KICKBAN) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") >= 2 Then
                    param1$ = getNextToken(szLine, " ") '#channel
                    param2$ = getNextToken(szLine, " ") '#channel
                    Call SendData(mvarlpIRC_SOCKET, "MODE " & param1$ & " +b " & param2$ & vbCrLf)
                    Call SendData(mvarlpIRC_SOCKET, "KICK " & param1$ & " " & param2$ & " :" & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.kickban #channel nickname kickmessage>"
                End If
                
            Case ".permban"
                BotConns.Item(szKey).LastCommand = szCommand 'priv_permban
                
            Case ".say"
                BotConns.Item(szKey).LastCommand = szCommand
                If BotConns.Item(szKey).CheckBotPerms(PRIV_SAY) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") >= 2 Then
                    param1$ = getNextToken(szLine, " ") 'target
                    Call SendData(mvarlpIRC_SOCKET, "PRIVMSG " & param1$ & " :" & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.say TARGET my message goes here>"
                End If
            
            Case ".who"
                BotConns.Item(szKey).LastCommand = szCommand
                Call SendData(BotConns.Item(szKey).ConnID, "- Begin Who List -" & vbCrLf)
                For num = 1 To BotConns.count
                    If BotConns.Item(num).Status = 2 Then
                        Call SendData(BotConns.Item(szKey).ConnID, "- " & BotConns.Item(num).User & vbCrLf)
                    End If
                Next num
                Call SendData(BotConns.Item(szKey).ConnID, "- End of Who List -" & vbCrLf)
            
            Case ".topic"
                BotConns.Item(szKey).LastCommand = szCommand
                If BotConns.Item(szKey).CheckChanPerms(GetChannel(szLine), PRIV_TOPIC) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") >= 2 Then
                    param1$ = getNextToken(szLine, " ") '#channel
                    Call SendData(mvarlpIRC_SOCKET, "TOPIC " & param1$ & " :" & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.topic #channel my topic goes here>"
                End If
                
            Case ".join"
                BotConns.Item(szKey).LastCommand = szCommand
                If BotConns.Item(szKey).CheckBotPerms(PRIV_JOIN) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") = 1 Then
                    Call SendData(mvarlpIRC_SOCKET, "JOIN :" & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.join #channel>"
                End If
                
            Case ".part"
                BotConns.Item(szKey).LastCommand = szCommand
                If BotConns.Item(szKey).CheckBotPerms(PRIV_PART) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") = 1 Then
                    Call SendData(mvarlpIRC_SOCKET, "PART :" & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.part #channel>"
                End If
                
            Case ".passwd"
                BotConns.Item(szKey).LastCommand = szCommand
                If CountTokens(szLine, " ") = 1 Then
                    param1$ = getNextToken(szLine, " ") 'password
                    WritePrivProfString Encrypt(param1$), BotConns.Item(szKey).User, "pass", "pass.txt"
                    Call SendData(BotConns.Item(szKey).ConnID, "-Password for " & BotConns.Item(szKey).User & " has been changed!-" & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.passwd new-password-goes-here>"
                End If
                
            Case ".adduser"
                BotConns.Item(szKey).LastCommand = szCommand 'priv_adduser
                If BotConns.Item(szKey).CheckBotPerms(PRIV_ADDUSER) Then
                    SendBotError ConnID, "Wrong Access Privileges!"
                    Exit Sub
                End If
                If CountTokens(szLine, " ") = 2 Then
                    'Call SendData(mvarlpIRC_SOCKET, "PART :" & szLine & vbCrLf)
                Else
                    SendBotError ConnID, "Syntax Error! <.adduser username password>"
                End If
                
            Case ".deluser"
                BotConns.Item(szKey).LastCommand = szCommand 'priv_deluser
                
            Case ".lsusers"
                BotConns.Item(szKey).LastCommand = szCommand 'priv_lsusers
                
            Case ".info"
                BotConns.Item(szKey).LastCommand = szCommand
                
            Case ".note"
                BotConns.Item(szKey).LastCommand = szCommand
                
            Case ".help"
                BotConns.Item(szKey).LastCommand = szCommand
                If szLine = szCommand Then
                    Open App.Path & "\bothelp.txt" For Input As #1
                        param1$ = Input(LOF(1), 1)
                    Close #1
                    Call SendData(ConnID, "- Begin Help File -" & vbCrLf)
                    Call SendData(ConnID, param1$ & vbCrLf)
                    Call SendData(ConnID, "- End Help File -" & vbCrLf)
                End If
                
            Case Else
                If Mid(szCommand, 1, 1) = "." Then
                    BotConns.Item(szKey).LastCommand = szCommand
                    'hit our script
                    'pass it command, params
                Else
                    'msg to party line
                    If LCase$(szCommand) = LCase$(szLine) Then
                        SendToPartyLine "<" & BotConns.Item(szKey).User & "> " & szLine
                    Else
                        SendToPartyLine "<" & BotConns.Item(szKey).User & "> " & szCommand & " " & szLine
                    End If
                End If
        End Select
    Else 'it was false, so everything goes to the script
        If Mid(szCommand, 1, 1) = "." Then
            BotConns.Item(szKey).LastCommand = szCommand
            'hit our script
            'pass it command, params
        Else
            'msg to party line
            If LCase$(szCommand) = LCase$(szLine) Then
                SendToPartyLine "<" & BotConns.Item(szKey).User & "> " & szLine
            Else
                SendToPartyLine "<" & BotConns.Item(szKey).User & "> " & szCommand & " " & szLine
            End If
        End If
    End If
End Sub

Public Sub SendToPartyLine(ByVal szLine As String)
    Dim num As Integer
    For num = 1 To BotConns.count
        If BotConns.Item(num).Status = 2 Then
            Call SendData(BotConns.Item(num).ConnID, szLine & vbCrLf)
        End If
    Next num
End Sub

Private Sub SendBotError(ByVal ConnID As Long, ByVal szErr As String)
    Call SendData(ConnID, "ERROR: " & szErr & vbCrLf)
End Sub

Public Function GetVersion() As String
    GetVersion = "[" & App.Title & "] " & App.Major & "." & App.Minor & "." & App.Revision
End Function
Private Sub SetLevels()
    PRIV_OP = CInt(GetPrivProfString("Levels", "op", "bot.ini"))
    PRIV_DEOP = CInt(GetPrivProfString("Levels", "deop", "bot.ini"))
    PRIV_PV = CInt(GetPrivProfString("Levels", "pv", "bot.ini"))
    PRIV_MV = CInt(GetPrivProfString("Levels", "mv", "bot.ini"))
    PRIV_KICK = CInt(GetPrivProfString("Levels", "kick", "bot.ini"))
    PRIV_BAN = CInt(GetPrivProfString("Levels", "ban", "bot.ini"))
    PRIV_UNBAN = CInt(GetPrivProfString("Levels", "unban", "bot.ini"))
    PRIV_SAY = CInt(GetPrivProfString("Levels", "say", "bot.ini"))
    PRIV_KICKBAN = CInt(GetPrivProfString("Levels", "kickban", "bot.ini"))
    PRIV_PERMBAN = CInt(GetPrivProfString("Levels", "permban", "bot.ini"))
    PRIV_TOPIC = CInt(GetPrivProfString("Levels", "topic", "bot.ini"))
    PRIV_JOIN = CInt(GetPrivProfString("Levels", "join", "bot.ini"))
    PRIV_PART = CInt(GetPrivProfString("Levels", "part", "bot.ini"))
    PRIV_ADDUSER = CInt(GetPrivProfString("Levels", "adduser", "bot.ini"))
    PRIV_DELUSER = CInt(GetPrivProfString("Levels", "deluser", "bot.ini"))
    PRIV_LSUSERS = CInt(GetPrivProfString("Levels", "lsusers", "bot.ini"))
    PRIV_SETPERMS = CInt(GetPrivProfString("Levels", "setperms", "bot.ini"))
End Sub


Private Sub Class_Initialize()
    SetLevels
End Sub


