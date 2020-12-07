VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CClientManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************************************
'   Client Manager
'       Author: Nicholas J. Felmlee
'       Email: Nick@Felmlee.Com
'       Last Revision: 7/18/98
'
'       Internal Dependencies:
'           WskSock.bas
'           CScriptManager.cls
'           CStringStack.cls
'           CClientManager.cls
'
'       External Dependencies:
'           BotObject.dll
'           DCCObject.dll
'           IdentdObject.dll
'           IRCObject.dll
'
'**************************************************************

Option Explicit
'local variable(s) to hold property value(s)
Private mvarUsername As String
Private mvarRealName As String
Private mvarServerName As String
Private mvarServerPort As Integer
Private mvarEmail As String
Private mvarUserInfo As String
Private mvarPingReply As String
Private mvarQuitMsg As String 'local copy
Private mvarFinger As String
Private mvarNickName As String
Private mvarAltNickName As String
Private mvarHighLight As String
Private mvarRetries As Integer
Private mvarInvisible As Boolean
Private mvarIdentdName As String
Private mvarIdentdSystem As String
Private mvarCLR_MESSAGE As Integer
Private mvarCLR_MODE As Integer
Private mvarCLR_KICK As Integer
Private mvarCLR_CTCP As Integer
Private mvarCLR_JOIN As Integer
Private mvarCLR_PART As Integer
Private mvarCLR_ACTION As Integer
Private mvarCLR_NICK As Integer
Private mvarCLR_NICKBACK As Integer
Private mvarCLR_TEXTBACK As Integer
Private mvarCLR_HIGHLIGHT As Integer
Private mvarCLR_NOTICE As Integer
Private mvarFontName As String
Private mvarFontSize As Integer
Private mvarFontBold As Boolean
Private mvarDCC_AutoGet As Boolean
Private mvarDCC_SaveDir As String
Private mvarDCC_ChunkSize As Integer
Private mvarDCC_TimeOut As Integer
Private mvarBOT_DisableDefaultOps As Boolean
Private mvarChanWins As Collection
Private mvarMessWins As Collection
Private mvarDccWins As Collection
Private mvarhWnd As Long
'Constants
Const CLIENTINFO = "BoDeBoT II" ' (" & App.Major & "." & App.Minor & ") - Not just a BoT, I'm also a Client"
Const VERSION = "BoDeBoT II " '(" & App.Major & "." & App.Minor & ") - Not just a BoT, I'm also a Client"
'Objects
Private WithEvents IRC_Ob As CIRCFactory
Attribute IRC_Ob.VB_VarHelpID = -1
Private WithEvents ID_Ob As CIdentdFactory
Attribute ID_Ob.VB_VarHelpID = -1
Private WithEvents BOT_Ob As CBotFactory
Attribute BOT_Ob.VB_VarHelpID = -1
Private WithEvents DCC_Ob As CDCCFactory
Attribute DCC_Ob.VB_VarHelpID = -1
Private SCR_Ob As CScriptManager
Private HookMan As CHookManager
'just plain variables (not properties)
Public Connected As Boolean
Public NOTIFYNICKS As String
'our events
Public Event onDisconnect(ByVal szHost As String) 'fires when irc connection drops
Public Event onConnect(ByVal szHost As String) 'fires when we connect to irc
Public Event onNewChannelWindow(ByVal szChan As String) 'fires when we join a channel
Public Event onNewMessageWindow(ByVal szNick As String) 'fies when we get a new priv msg
Public Event onKillChannelWindow(ByVal szChan As String) 'fires when a channel window should be closed
Public Event onKillMessageWindow(ByVal szNick As String) 'fires when a message window should be closed
Public Event onUpdateNotifyWindow(ByVal szNicks As String) 'fires when we get an ISON rply
Public Event onBotConnection(ByVal szNick As String, ByVal dwCount As Integer) 'fires when we get a bot connection
Public Event onBotDisconnection(ByVal szNick As String, ByVal dwCount As Integer) 'fires when a user leaves the bot
Public Event onBotSendToScript(UserOb As Object, ByVal szCommand As String, ByVal szParams As String)
Public Event onQueryAppVersion(ByVal szInfo As String) 'fires when we call GetVersion() returns versions of components
Public Event onSelectFileForDCCSend(ByVal szNick As String) 'fires when we want to dcc send (/dcc) we need to have the user select a file at this point
Public Event onHighLightEvent(ByVal szChannel As String)

Public Property Let QuitMsg(ByVal vData As String)
    mvarQuitMsg = vData
End Property

Public Property Get QuitMsg() As String
    QuitMsg = mvarQuitMsg
End Property

Public Property Set DccWins(ByVal vData As Object)
    Set mvarDccWins = vData
End Property


Public Property Get DccWins() As Collection
    Set DccWins = mvarDccWins
End Property

Public Property Let hWnd(ByVal vData As Long)
    mvarhWnd = vData
End Property
Public Property Get hWnd() As Long
    hWnd = mvarhWnd
End Property
Public Property Set MessWins(ByVal vData As Object)
    Set mvarMessWins = vData
End Property


Public Property Get MessWins() As Collection
    Set MessWins = mvarMessWins
End Property


Public Property Set ChanWins(ByVal vData As Object)
    Set mvarChanWins = vData
End Property


Public Property Get ChanWins() As Collection
    Set ChanWins = mvarChanWins
End Property

Public Property Let BOT_DisableDefaultOps(ByVal vData As Boolean)
    mvarBOT_DisableDefaultOps = vData
End Property


Public Property Get BOT_DisableDefaultOps() As Boolean
    BOT_DisableDefaultOps = mvarBOT_DisableDefaultOps
End Property

Public Property Let DCC_TimeOut(ByVal vData As Integer)
    mvarDCC_TimeOut = vData
End Property


Public Property Get DCC_TimeOut() As Integer
    DCC_TimeOut = mvarDCC_TimeOut
End Property



Public Property Let DCC_ChunkSize(ByVal vData As Integer)
    mvarDCC_ChunkSize = vData
End Property


Public Property Get DCC_ChunkSize() As Integer
    DCC_ChunkSize = mvarDCC_ChunkSize
End Property



Public Property Let DCC_SaveDir(ByVal vData As String)
    mvarDCC_SaveDir = vData
End Property


Public Property Get DCC_SaveDir() As String
    DCC_SaveDir = mvarDCC_SaveDir
End Property



Public Property Let DCC_AutoGet(ByVal vData As Boolean)
    mvarDCC_AutoGet = vData
End Property


Public Property Get DCC_AutoGet() As Boolean
    DCC_AutoGet = mvarDCC_AutoGet
End Property



Public Property Let FontBold(ByVal vData As Boolean)
    mvarFontBold = vData
End Property


Public Property Get FontBold() As Boolean
    FontBold = mvarFontBold
End Property



Public Property Let FontSize(ByVal vData As Integer)
    mvarFontSize = vData
End Property


Public Property Get FontSize() As Integer
    FontSize = mvarFontSize
End Property



Public Property Let FontName(ByVal vData As String)
    mvarFontName = vData
End Property


Public Property Get FontName() As String
    FontName = mvarFontName
End Property



Public Property Let CLR_NOTICE(ByVal vData As Integer)
    mvarCLR_NOTICE = vData
End Property


Public Property Get CLR_NOTICE() As Integer
    CLR_NOTICE = mvarCLR_NOTICE
End Property



Public Property Let CLR_HIGHLIGHT(ByVal vData As Integer)
    mvarCLR_HIGHLIGHT = vData
End Property


Public Property Get CLR_HIGHLIGHT() As Integer
    CLR_HIGHLIGHT = mvarCLR_HIGHLIGHT
End Property



Public Property Let CLR_TEXTBACK(ByVal vData As Integer)
    mvarCLR_TEXTBACK = vData
End Property


Public Property Get CLR_TEXTBACK() As Integer
    CLR_TEXTBACK = mvarCLR_TEXTBACK
End Property



Public Property Let CLR_NICKBACK(ByVal vData As Integer)
    mvarCLR_NICKBACK = vData
End Property


Public Property Get CLR_NICKBACK() As Integer
    CLR_NICKBACK = mvarCLR_NICKBACK
End Property



Public Property Let CLR_NICK(ByVal vData As Integer)
    mvarCLR_NICK = vData
End Property


Public Property Get CLR_NICK() As Integer
    CLR_NICK = mvarCLR_NICK
End Property



Public Property Let CLR_ACTION(ByVal vData As Integer)
    mvarCLR_ACTION = vData
End Property


Public Property Get CLR_ACTION() As Integer
    CLR_ACTION = mvarCLR_ACTION
End Property



Public Property Let CLR_PART(ByVal vData As Integer)
    mvarCLR_PART = vData
End Property


Public Property Get CLR_PART() As Integer
    CLR_PART = mvarCLR_PART
End Property



Public Property Let CLR_JOIN(ByVal vData As Integer)
    mvarCLR_JOIN = vData
End Property


Public Property Get CLR_JOIN() As Integer
    CLR_JOIN = mvarCLR_JOIN
End Property



Public Property Let CLR_CTCP(ByVal vData As Integer)
    mvarCLR_CTCP = vData
End Property


Public Property Get CLR_CTCP() As Integer
    CLR_CTCP = mvarCLR_CTCP
End Property



Public Property Let CLR_KICK(ByVal vData As Integer)
    mvarCLR_KICK = vData
End Property


Public Property Get CLR_KICK() As Integer
    CLR_KICK = mvarCLR_KICK
End Property



Public Property Let CLR_MODE(ByVal vData As Integer)
    mvarCLR_MODE = vData
End Property


Public Property Get CLR_MODE() As Integer
    CLR_MODE = mvarCLR_MODE
End Property



Public Property Let CLR_MESSAGE(ByVal vData As Integer)
    mvarCLR_MESSAGE = vData
End Property


Public Property Get CLR_MESSAGE() As Integer
    CLR_MESSAGE = mvarCLR_MESSAGE
End Property



Public Property Let IdentdSystem(ByVal vData As String)
    mvarIdentdSystem = vData
End Property


Public Property Get IdentdSystem() As String
    IdentdSystem = mvarIdentdSystem
End Property



Public Property Let IdentdName(ByVal vData As String)
    mvarIdentdName = vData
End Property


Public Property Get IdentdName() As String
    IdentdName = mvarIdentdName
End Property



Public Property Let Invisible(ByVal vData As Boolean)
    mvarInvisible = vData
End Property


Public Property Get Invisible() As Boolean
    Invisible = mvarInvisible
End Property



Public Property Let Retries(ByVal vData As Integer)
    mvarRetries = vData
End Property


Public Property Get Retries() As Integer
    Retries = mvarRetries
End Property



Public Property Let HighLight(ByVal vData As String)
    mvarHighLight = vData
End Property


Public Property Get HighLight() As String
    HighLight = mvarHighLight
End Property



Public Property Let AltNickName(ByVal vData As String)
    mvarAltNickName = vData
End Property


Public Property Get AltNickName() As String
    AltNickName = mvarAltNickName
End Property



Public Property Let NickName(ByVal vData As String)
    mvarNickName = vData
End Property


Public Property Get NickName() As String
    NickName = mvarNickName
End Property



Public Property Let Finger(ByVal vData As String)
    mvarFinger = vData
End Property


Public Property Get Finger() As String
    Finger = mvarFinger
End Property



Public Property Let PingReply(ByVal vData As String)
    mvarPingReply = vData
End Property


Public Property Get PingReply() As String
    PingReply = mvarPingReply
End Property



Public Property Let UserInfo(ByVal vData As String)
    mvarUserInfo = vData
End Property


Public Property Get UserInfo() As String
    UserInfo = mvarUserInfo
End Property



Public Property Let Email(ByVal vData As String)
    mvarEmail = vData
End Property


Public Property Get Email() As String
    Email = mvarEmail
End Property



Public Property Let ServerPort(ByVal vData As Integer)
    mvarServerPort = vData
End Property


Public Property Get ServerPort() As Integer
    ServerPort = mvarServerPort
End Property



Public Property Let ServerName(ByVal vData As String)
    mvarServerName = vData
End Property


Public Property Get ServerName() As String
    ServerName = mvarServerName
End Property



Public Property Let RealName(ByVal vData As String)
    mvarRealName = vData
End Property


Public Property Get RealName() As String
    RealName = mvarRealName
End Property



Public Property Let Username(ByVal vData As String)
    mvarUsername = vData
End Property


Public Property Get Username() As String
    Username = mvarUsername
End Property

Private Sub BOT_Ob_onUserCommand(ByVal UserOb As Object, ByVal szCommand As String, ByVal szParams As String)
    'update last action in our bot form
    BotSetup.ListView1.ListItems(UserOb.User).SubItems(1) = "<" & szCommand & "> " & szParams
    SCR_Ob.onBotCommands UserOb, szCommand, szParams
    'TO DO: send to bot log
End Sub

Private Sub BOT_Ob_onUserJoin(ByVal szUser As String, ByVal szHost As String, ByVal dwCount As Integer)
    RaiseEvent onBotConnection(szUser, dwCount)
    BotSetup.ListView1.ListItems.Add , szUser, szUser
    'TO DO: send to bot log
End Sub

Private Sub BOT_Ob_onUserPart(ByVal szName As String, ByVal dwCount As Integer)
    RaiseEvent onBotDisconnection(szName, dwCount)
    BotSetup.ListView1.ListItems.Remove szName
    'TO DO: send to bot log
End Sub

Private Sub Class_Initialize()
    'we need a hook manager before we do anything
    Set HookMan = New CHookManager
    'init the connected switch
    Connected = False
    'Lets fire up winsock
    StartWinsock ("")
    'set our collections
    Set Me.ChanWins = New Collection
    Set Me.MessWins = New Collection
    Set Me.DccWins = New Collection
    'read ini
    GetStoredConfig
    'start our servers
    StartBotServer
    StartIdentServer
    StartDCCFactory
    Set SCR_Ob = New CScriptManager
    SCR_Ob.AddObject Me, "BOT"
End Sub

Private Sub Class_Terminate()
    'MsgBox "DIE MOTHER FUCKER!"
    'kill our hookmanager
    'note that the hook manager will unhook all on its onTerminate
    Set HookMan = Nothing
    Set IRC_Ob = Nothing
    Set ID_Ob = Nothing
    Set SCR_Ob = Nothing
    Set DCC_Ob = Nothing
    Set BOT_Ob = Nothing
    'kill winsock
    EndWinsock
End Sub

Private Sub StartIdentServer()
'STARTS THE IDENTD SERVER
    Set ID_Ob = New CIdentdFactory
    Load IDSRV 'load invisible form so we can hook
    ID_Ob.hWnd = IDSRV.hWnd
    'add object to the hookmanager
    HookMan.AddHook ID_Ob, 1025
    'init some of the properties
    ID_Ob.IdentName = mvarIdentdName
    ID_Ob.IdentSystem = mvarIdentdSystem
    ID_Ob.IdentPort = 113
    'start the identd server
    ID_Ob.Start
End Sub

Private Sub StartBotServer()
'STARTS THE BOT SERVER
    Set BOT_Ob = New CBotFactory
    Load BotSetup 'load form so we can hook
    BOT_Ob.hWnd = BotSetup.hWnd
    'add the object to the hookmanager
    HookMan.AddHook BOT_Ob, 1025
End Sub
Private Sub StartDCCFactory()
'STARTS OUR DCC FACTORY
    Set DCC_Ob = New CDCCFactory
End Sub

Public Sub ConnectToIrc()
'CREATES A NEW IRCFACTORY OBJECT AND
'CONNECTS US TO IRC. ONLY WILL DO THIS IF
'WE ARE NOT CURRENTLY CONNECTED.
    'if we're already connected then
    'lets get the hell out of this sub()
    If Connected = True Then Exit Sub
        'first close any open channel windows
        CloseAllChannels
        'next lets create a new IRCFactory object
        If Not (IRC_Ob Is Nothing) Then
            HookMan.DelHook IRC_Ob, 1025
            Set IRC_Ob = Nothing
        End If
        Set IRC_Ob = New CIRCFactory
        IRC_Ob.hWnd = mvarhWnd
        HookMan.AddHook IRC_Ob, 1025
        'load our props from the ini again
        GetStoredConfig
        'now init some stuff in the factory
        With IRC_Ob
            .Connected = False
            .Port = mvarServerPort
            .ServerName = mvarServerName
            .NickName = mvarNickName
            .RealName = mvarRealName
            .Retries = mvarRetries
            .Username = mvarUsername
            .ServerName = mvarServerName
        End With
        'display current build
        GetVersion
        'now connect
        IRC_Ob.Connect
End Sub

Public Sub DisconnectFromIrc()
'DISCONNECTS US FROM IRC AND CLEANS UP
'THE IRC FACTORY. ONLY WILL DO THIS IF
'WE ARE NOT ALREADY DISCONNECTED.
    'If we are already disconnected then
    'lets get the hell out of here!
    If Connected = False Then Exit Sub
    'we'll start by closing any open channel windows
    CloseAllChannels
    'send irc a quit message so we don't
    'end up with a ghost.
    IRC_Ob.Send "QUIT :" & mvarQuitMsg & vbCrLf
    'next have the object close the socket just in case
    IRC_Ob.Disconnect
    'fire out disconnect event
    RaiseEvent onDisconnect(IRC_Ob.ServerName)
    'remove the irc factory from the hookmanager
    HookMan.DelHook IRC_Ob, 1025
    'cleanup some memory
    Set IRC_Ob = Nothing
    Connected = False
End Sub
Private Sub GetStoredConfig()
    Dim di As Long
    Dim retval As String * 4096
    '[CLIENT] Section
    mvarUsername = GetFromIni("client", "username")
    mvarRealName = GetFromIni("client", "realname")
    mvarEmail = GetFromIni("client", "email")
    mvarUserInfo = GetFromIni("client", "userinfo")
    mvarPingReply = GetFromIni("client", "pingreply")
    mvarFinger = GetFromIni("client", "finger")
    mvarQuitMsg = GetFromIni("client", "quit")
    mvarRetries = CInt(GetFromIni("client", "retries"))
    mvarNickName = GetFromIni("client", "nickname")
    mvarAltNickName = GetFromIni("client", "altnickname")
    mvarHighLight = GetFromIni("client", "highlight")
    mvarInvisible = CBool(GetFromIni("client", "invisible"))
    
    '[IDENTD]
    mvarIdentdName = GetFromIni("identd", "identname")
    mvarIdentdSystem = GetFromIni("identd", "identsys")
    
    '[DCC]
    mvarDCC_AutoGet = CBool(GetFromIni("dcc", "dccautoget"))
    mvarDCC_SaveDir = GetFromIni("dcc", "dccsavedir")
    mvarDCC_ChunkSize = CInt(GetFromIni("dcc", "dccpacketsize"))
    mvarDCC_TimeOut = CInt(GetFromIni("dcc", "dcctimeout"))
    
    '[COLOR] SECTION
    mvarCLR_MODE = CInt(GetFromIni("color", "modecolor"))
    mvarCLR_KICK = CInt(GetFromIni("color", "kickcolor"))
    mvarCLR_CTCP = CInt(GetFromIni("color", "ctcpcolor"))
    mvarCLR_JOIN = CInt(GetFromIni("color", "joincolor"))
    mvarCLR_PART = CInt(GetFromIni("color", "partcolor"))
    mvarCLR_ACTION = CInt(GetFromIni("color", "actioncolor"))
    mvarCLR_NICK = CInt(GetFromIni("color", "nickcolor"))
    mvarCLR_NICKBACK = CInt(GetFromIni("color", "nickbackcolor"))
    mvarCLR_TEXTBACK = CInt(GetFromIni("color", "textbackcolor"))
    mvarCLR_MESSAGE = CInt(GetFromIni("color", "messagecolor"))
    mvarCLR_HIGHLIGHT = CInt(GetFromIni("color", "highlightcolor"))
    mvarCLR_NOTICE = CInt(GetFromIni("color", "noticecolor"))
    
    '[FONT] section
    mvarFontName = GetFromIni("font", "fontname")
    mvarFontSize = CInt(GetFromIni("font", "fontsize"))
    mvarFontBold = CBool(GetFromIni("font", "fontbold"))
    
    mvarServerName = GetFromIni("client", "servername")
    mvarServerPort = 6667
    '[BOT]
    mvarBOT_DisableDefaultOps = CBool(GetFromIni("bot", "bypassexe"))
End Sub
Private Sub CloseAllChannels()
    While mvarChanWins.count > 0
        Unload mvarChanWins.Item(1)
    Wend
End Sub
Private Sub CloseAllMessages()
    While mvarMessWins.count > 0
        Unload mvarMessWins.Item(1)
    Wend
End Sub
Private Sub CloseAllDCCs()
    While mvarDccWins.count > 0
        Unload mvarDccWins.Item(1)
    Wend
End Sub
Public Sub CleanHouse()
    'CLEANS UP ANY GUI FORMS WE HAVE STORED IN
    'OUR FORM COLLECTIONS. THIS NEEDS TO BE DONE
    'PRIOR TO DESTROYING THIS CLASS
    CloseAllChannels
    CloseAllMessages
    CloseAllDCCs
End Sub
Private Sub ChanNew(SChannel As String)
'CREATES A NEW CHANNEL WINDOW
'AND ADDS IT TO OUR CHANNEL WINDOWS
'COLLECTION.
    Dim ChanWin As New chan
    ChanWin.Caption = "[" & SChannel & "]"
    mvarChanWins.Add ChanWin, SChannel
    ChanWin.Tag = SChannel
    RaiseEvent onNewChannelWindow(SChannel)
    ChanWin.Show
End Sub
Private Sub MessNew(the_nick As String)
'CREATES A NEW MESSAGE WINDOW AND ADDS
'IT TO OUR MESSAGE WINDOW COLLECTION.
    Dim MessWin As New messg
    MessWin.Caption = the_nick
    MessWin.Tag = the_nick
    mvarMessWins.Add MessWin, the_nick
    RaiseEvent onNewMessageWindow(the_nick)
End Sub

Private Sub KickSomebody(chan As String, nick As String, kkmessage As String)
    IRC_Ob.Send "KICK " & chan & " " & nick & " :" & kkmessage & vbCrLf
End Sub

Private Sub SendCTCP(nick As String, command As String, Reply As String)
    IRC_Ob.Send "NOTICE " & nick & " :" & Chr$(1) & UCase$(command) & " " & Reply & Chr$(1) & vbCrLf
End Sub

Public Sub ProcessUserCommands(thebox As Control, ByVal szLine As String, Optional ByVal szTarget)
        Dim szCommand$, temp$, param1$, param2$, param3$
        szCommand = getNextToken(szLine, " ")
        Select Case LCase$(szCommand)
            Case "/join"
                IRC_Ob.Send "JOIN :" & szLine & vbCrLf
                UpdateRTB thebox, "**" & NickName & "** ->JOIN " & szLine, mvarCLR_JOIN
            
            Case "/msg"
                param1$ = getNextToken(szLine, " ")
                IRC_Ob.Send "PRIVMSG " & param1$ & " :" & szLine & vbCrLf
                UpdateRTB thebox, "[PRIVATE] **" & param1$ & "**   " & szLine, mvarCLR_MESSAGE
                                
            Case "/whois"
                IRC_Ob.Send "WHOIS :" & szLine & vbCrLf
            
            Case "/nick"
                IRC_Ob.Send "NICK :" & Trim(szLine) & vbCrLf
                
            Case "/away"
                IRC_Ob.Send "AWAY :" & szLine & vbCrLf
                                    
            Case "/whowas"
                IRC_Ob.Send "WHOWAS :" & szLine & vbCrLf
            
            Case "/who"
                IRC_Ob.Send "WHO :" & szLine & vbCrLf
            
            Case "/mode"
                IRC_Ob.Send "MODE " & szLine & vbCrLf
                UpdateRTB thebox, "[Mode]** " & szLine, mvarCLR_MODE
                
            Case "/topic"
                IRC_Ob.Send "TOPIC " & getNextToken(szLine, " ") & " :" & szLine & vbCrLf
                UpdateRTB thebox, "[TOPIC]** " & szLine, mvarCLR_MODE
            Case "/ctcp"
                param1$ = getNextToken(szLine, " ")
                IRC_Ob.Send "PRIVMSG " & param1$ & " :" & Chr(1) & szLine & Chr(1) & vbCrLf
                UpdateRTB thebox, " **" & param1$ & "**   " & szLine, mvarCLR_CTCP
            
            Case "/notice"
                param1$ = getNextToken(szLine, " ")
                IRC_Ob.Send "NOTICE " & param1$ & " :" & szLine & vbCrLf
                UpdateRTB thebox, "[NOTICE]**" & param1$ & "**   " & szLine, mvarCLR_CTCP
            Case "/dcc"
                RaiseEvent onSelectFileForDCCSend(szLine)
                
            Case "/action"
                IRC_Ob.Send "PRIVMSG " & szTarget & " :" & Chr(1) & "ACTION" & " " & szLine & Chr(1) & vbCrLf
                UpdateRTB thebox, "*" & mvarNickName & " " & szLine, mvarCLR_ACTION
            
                        
            Case "/chat"
                Dim schat As DCCChat
                Set schat = New DCCChat
                Load schat
                param1$ = CStr(schat.hWnd)
                DccWins.Add schat, param1$
                schat.Show
                DCC_Ob.Sendchat szLine, schat
                
            Case "/invite"
                IRC_Ob.Send "INVITE " & szLine & vbCrLf
                
            Case "/notify"
                If InStr(1, Trim(szLine), " ") <> 0 Then
                    Do Until LCase$(param1$) = LCase$(szLine)
                        param1$ = getNextToken(szLine, " ")
                        If InStr(1, LCase$(NOTIFYNICKS), LCase$(param1$)) = 0 Then
                            NOTIFYNICKS = NOTIFYNICKS & " " & param1$
                        End If
                    Loop
                Else
                    If InStr(1, LCase$(NOTIFYNICKS), LCase$(Trim(szLine))) = 0 Then
                        NOTIFYNICKS = NOTIFYNICKS & " " & szLine
                    End If
                End If
                IRC_Ob.Send "ISON " & NOTIFYNICKS & vbCrLf

            Case "/server"
                If Connected = True Then
                    DisconnectFromIrc
                End If
                mvarServerName = szLine
                ConnectToIrc
            Case Else
                If Not Mid(szCommand, 1, 1) = "/" Then
                'send to target
                    If Not szTarget = "" Then
                        If Not (szCommand = szLine) Then
                            IRC_Ob.Send "PRIVMSG " & szTarget & " :" & szCommand & " " & szLine & vbCrLf
                        Else
                            IRC_Ob.Send "PRIVMSG " & szTarget & " :" & szLine & vbCrLf
                        End If
                    End If
                Else
                'send to script
                    If Mid(szTarget, 1, 1) = "#" Then
                        SCR_Ob.onChanCommands mvarChanWins.Item(szTarget), szCommand, szLine
                    Else
                        SCR_Ob.onMsgCommands mvarMessWins.Item(szTarget), szCommand, szLine
                    End If
                End If
        End Select
End Sub

Private Sub DCC_Ob_onAddDCCConn(Ob As Object)
    HookMan.AddHook Ob, 1025
End Sub

Private Sub DCC_Ob_onDCCChatClose(ByVal szNick As String, targfrm As Object)
    targfrm.Caption = targfrm.Caption & " [Session Closed]"
    targfrm.msendrtb.Locked = True
End Sub

Private Sub DCC_Ob_onDCCFileDone(ByVal szNick As String, ByVal szFileName As String, ByVal szHostAddr As String, ByVal lpFileSize As Long, targfrm As Object)
    targfrm.ProgressBar1.Value = targfrm.ProgressBar1.Max
    targfrm.Caption = targfrm.Caption & " [Completed]"
    targfrm.Timer1.Enabled = False
End Sub

Private Sub DCC_Ob_onDCCFileStart(ByVal szNick As String, ByVal szFileName As String, ByVal szHostAddr As String, ByVal lpFileSize As Long, ByVal chunksize As Long, targfrm As Object, ByVal szOp As String)
    targfrm.ProgressBar1.Min = 0
    targfrm.Caption = "[" & szNick & "] DCC " & szOp
    targfrm.ProgressBar1.Max = CLng(lpFileSize)
    targfrm.Label1.Caption = szHostAddr
    targfrm.Label2.Caption = ReturnFileName(szFileName)
    If lpFileSize >= 1000000 Then
        targfrm.Label3.Caption = CLng(CLng(lpFileSize) / 1000000) & " Megabytes"
    ElseIf lpFileSize >= 1000 Then
        targfrm.Label3.Caption = CLng(CLng(lpFileSize) / 1000) & " Kilobytes"
    Else
        targfrm.Label3.Caption = CLng(lpFileSize) & " Bytes"
    End If
    If Not mvarDCC_TimeOut < 60 Then
        mvarDCC_TimeOut = 30
    End If
    targfrm.Timer1.Interval = 1000 * mvarDCC_TimeOut
    targfrm.Timer1.Enabled = True
End Sub

Private Sub DCC_Ob_onDelDCCConn(Ob As Object)
    HookMan.DelHook Ob, 1025
End Sub


Private Sub DCC_Ob_onUpdateProgress(ByVal lpChunk As Long, targfrm As Object)
    If targfrm.ProgressBar1.Value + 1 < targfrm.ProgressBar1.Max Then
        targfrm.ProgressBar1.Value = targfrm.ProgressBar1.Value + lpChunk
    End If
    targfrm.Timer1.Enabled = False
    targfrm.Timer1.Enabled = True
End Sub

Private Sub DCC_Ob_onWindowOutput(ByVal szLine As String, targfrm As Object)
    UpdateRTB targfrm.messgrtb, szLine, mvarCLR_MESSAGE
End Sub


Private Sub IRC_Ob_onBeginNickList(ByVal szChan As String, ByVal szNicks As String)
    Dim temp As String
    Dim key As String
    Do Until temp = szNicks
        temp = getNextToken(szNicks, " ")
        If (InStr(1, temp, "@") <> 0 Or InStr(1, temp, "+") <> 0) Then
            key = Mid(temp, 2)
        Else
            key = temp
        End If
        'add the if in case we get a 353 before doing a NAMES
        If existsInList(mvarChanWins.Item(szChan).nicklist, key) = False Then
            mvarChanWins.Item(szChan).nicklist.ListItems.Add 1, key, temp
        End If
    Loop
End Sub

Private Sub IRC_Ob_onChanMode(ByVal szSetter As String, ByVal szChan As String, ByVal szMode As String, ByVal szParams As String)
    Debug.Print "MODE CHANGE! -> Target=" & szChan & " Mode=" & szMode
    UpdateRTB WConsole.RTB, "*** " & szSetter & " set " & szMode & " " & szParams & " on " & szChan, mvarCLR_MODE
    'get len of mode, if > 2 then loop to set each
    Dim sw As String
    Dim l As Integer, c As Integer
    Dim tempm As String
    Dim tempn As String
    sw = Mid(szMode, 1, 1) 'either + or -
    szMode = Mid(szMode, 2)
    tempn = szParams
    l = Len(szMode)
'    On Error GoTo kill
    For c = 0 To l
        tempm = Mid(szMode, 1, 1)
        tempn = getNextToken(tempn, " ")
        If l > 1 Then
            szMode = Mid(szMode, 2)
        End If
        Select Case UCase$(szMode)
            Case "O"
                If sw = "+" Then
                    mvarChanWins.Item(szChan).nicklist.ListItems(tempn).Text = "@" & tempn
                Else
                    mvarChanWins.Item(szChan).nicklist.ListItems(tempn).Text = tempn
                End If
            Case "B"
                If sw = "+" Then
                Else
                End If
            Case "V"
                If sw = "+" Then
                    mvarChanWins.Item(szChan).nicklist.ListItems(tempn).Text = "+" & tempn
                Else
                    mvarChanWins.Item(szChan).nicklist.ListItems(tempn).Text = tempn
                End If
            Case "I"
                If sw = "+" Then
                Else
                End If
            Case "T"
                If sw = "+" Then
                Else
                End If
            Case "N"
                If sw = "+" Then
                Else
                End If
            Case "L"
                If sw = "+" Then
                Else
                End If
            Case "M"
                If sw = "+" Then
                Else
                End If
            Case "P"
                If sw = "+" Then
                Else
                End If
            Case "S"
                If sw = "+" Then
                Else
                End If
            Case "K"
                If sw = "+" Then
                Else
                End If
        End Select
    Next c
End Sub

Private Sub IRC_Ob_onChanMsg(ByVal szSender As String, ByVal szChan As String, ByVal szMessage As String)
    If InStr(1, LCase(szMessage), LCase(HighLight)) <> 0 Then
        If Mid(szMessage, 1, 7) = Chr(1) & "ACTION" Then
            UpdateRTB mvarChanWins.Item(szChan).chanrtb, "*" & szSender & " " & Mid(szMessage, 8, Len(szMessage) - 8), mvarCLR_ACTION
            Call sndPlaySound(App.Path & "\bode2.wav", &H1)
            RaiseEvent onHighLightEvent(szChan)
        Else
            UpdateRTB mvarChanWins.Item(szChan).chanrtb, "<" & szSender & ">   " & szMessage, mvarCLR_HIGHLIGHT
            Call sndPlaySound(App.Path & "\bode2.wav", &H1)
            RaiseEvent onHighLightEvent(szChan)
        End If
    Else
        If Mid(szMessage, 1, 7) = Chr(1) & "ACTION" Then
            UpdateRTB mvarChanWins.Item(szChan).chanrtb, "*" & szSender & " " & Mid(szMessage, 8, Len(szMessage) - 8), mvarCLR_ACTION
        Else
            UpdateRTB mvarChanWins.Item(szChan).chanrtb, "<" & szSender & ">   " & szMessage, mvarCLR_MESSAGE
        End If
    End If
    SCR_Ob.onChannelMessage szSender, szChan, szMessage
End Sub

Private Sub IRC_Ob_onConnect(ByVal szHost As String)
    RaiseEvent onConnect(IRC_Ob.ServerName)
    Connected = True
    If Not (BOT_Ob Is Nothing) Then
        BOT_Ob.lpIRC_SOCKET = IRC_Ob.lpIRC_SOCKET
        DCC_Ob.lpIRC_SOCKET = IRC_Ob.lpIRC_SOCKET
    End If
End Sub

Private Sub IRC_Ob_onCTCP(ByVal szNick As String, ByVal szCommand As String, ByVal szExtras As String)
    Dim szDaddr As String
    Dim szDport As String
    Dim szKey As String
    Debug.Print "CTCP: " & szNick & " hit with a [" & szCommand & "] " & szExtras
    Select Case UCase$(szCommand)
        Case "VERSION" & Chr(1)
            UpdateRTB WConsole.RTB, "[VERSION] " & szNick, mvarCLR_CTCP
            SendCTCP szNick, "VERSION", VERSION & " (v" & App.Major & "." & App.Minor & ")"
        Case "PING" & Chr(1)
            UpdateRTB WConsole.RTB, "[PING] " & szNick, mvarCLR_CTCP
            SendCTCP szNick, "PING", PingReply
        Case "CLIENTINFO" & Chr(1)
            UpdateRTB WConsole.RTB, "[CLIENTINFO] " & szNick, mvarCLR_CTCP
            SendCTCP szNick, "CLIENTINFO", CLIENTINFO
        Case "USERINFO" & Chr(1)
            UpdateRTB WConsole.RTB, "[USERINFO] " & szNick, mvarCLR_CTCP
            SendCTCP szNick, "USERINFO", UserInfo
        Case "TIME" & Chr(1)
            UpdateRTB WConsole.RTB, "[TIME] " & szNick, mvarCLR_CTCP
            SendCTCP szNick, "TIME", Now
        Case "FINGER" & Chr(1)
            UpdateRTB WConsole.RTB, "[FINGER] " & szNick, mvarCLR_CTCP
            SendCTCP szNick, "FINGER", Finger
        Case "DCC"
            Dim r As Integer
            Dim dccsw As String
            dccsw = getNextToken(szExtras, " ")
            If LCase$(dccsw) = "send" Then
                Dim szFname As String
                Dim szFsize As String
                szFname = getNextToken(szExtras, " ")
                szDaddr = DecodeLongIPAddr(getNextToken(szExtras, " "))
                szDport = getNextToken(szExtras, " ")
                szFsize = Mid(szExtras, 1, Len(szExtras) - 1)
                If MsgBox("Accept " & szFname & " (" & szFsize & " Bytes) From " & szNick & "?", vbYesNo, "BoDeBoT: DCC") = 6 Then
                    'okie we said yes, so lets connect to the guy
                    Dim fileconn As DCCSend
                    Set fileconn = New DCCSend
                    Load fileconn
                    szKey = CStr(fileconn.hWnd)
                    DccWins.Add fileconn, szKey
                    fileconn.Show
                    DCC_Ob.GetFile szNick, mvarDCC_SaveDir & "\" & szFname, szFsize, szDaddr, CInt(szDport), 1024, fileconn
                End If
            ElseIf LCase$(dccsw) = "chat" Then
                Call getNextToken(szExtras, " ") 'trash that extra 'chat'
                szDaddr = DecodeLongIPAddr(getNextToken(szExtras, " "))
                szDport = Mid(szExtras, 1, Len(szExtras) - 1)
                If MsgBox("Accept DCC Chat From " & szNick & "?", vbYesNo, "BoDeBoT: DCC Chat") = 6 Then
                    'okie we said yes so lets open a dcc chat
                    Dim chatconn As DCCChat
                    Set chatconn = New DCCChat
                    Load chatconn
                    szKey = CStr(chatconn.hWnd)
                    DccWins.Add chatconn, szKey
                    chatconn.Show
                    If Mid(szDaddr, 1, 1) = "-" Then
                        szDaddr = Mid(szDaddr, 2)
                    End If
                    DCC_Ob.GetChat szNick, szDaddr, CInt(szDport), chatconn
                End If
            End If
        Case "LOGIN" & Chr(1)
            BOT_Ob.DisableDefaultOps = False
            BOT_Ob.lpIRC_SOCKET = IRC_Ob.lpIRC_SOCKET
            BOT_Ob.MaxConnections = 25
            BOT_Ob.GetNewConnection szNick
        Case Else
            SCR_Ob.onCTCPCommand szNick, szCommand, szExtras
    End Select
End Sub

Private Sub IRC_Ob_onDisconnect(ByVal szHost As String)
    Connected = False
    RaiseEvent onDisconnect(IRC_Ob.ServerName)
    SCR_Ob.onConnectionStatusChange False
End Sub

Private Sub IRC_Ob_onEndMOTD()
    SCR_Ob.onConnectionStatusChange True
End Sub

Private Sub IRC_Ob_onError(ByVal lpErrCode As Long, ByVal szDescription As String)
    UpdateRTB WConsole.RTB, "ERROR " & CStr(lpErrCode) & ": " & szDescription, 4
    Dim temp As String
    Select Case lpErrCode
        Case 433 'ERR_NICKNAMEINUSECase 433
            temp = mvarNickName
            mvarNickName = mvarAltNickName
            mvarAltNickName = temp
            IRC_Ob.Send "NICK :" & mvarNickName & vbCrLf
    End Select
End Sub

Private Sub IRC_Ob_onNickChanged(ByVal szOriginalNick As String, ByVal szNewNick As String)
    UpdateRTB WConsole.RTB, "*** " & szOriginalNick & " is now known as " & szNewNick & "!", mvarCLR_NOTICE
    Dim num As Integer
    Dim temp As String
    If LCase$(szOriginalNick) = LCase$(NickName) Then
        AltNickName = szOriginalNick
        NickName = szNewNick
    End If
        For num = 1 To mvarChanWins.count
            temp = ""
            If existsInList(ChanWins.Item(num).nicklist, szOriginalNick) Then
                If Len(ChanWins.Item(num).nicklist.ListItems(szOriginalNick).Text) = Len(szOriginalNick) + 1 Then
                    temp = Mid(ChanWins.Item(num).nicklist.ListItems(szOriginalNick).Text, 1, 1)
                End If
                mvarChanWins.Item(num).nicklist.ListItems.Remove szOriginalNick
                mvarChanWins.Item(num).nicklist.ListItems.Add 1, szNewNick, temp & szNewNick
            End If
            'num = num + 1
        Next num
End Sub


Private Sub IRC_Ob_onNickInvite(ByVal szNick As String, ByVal szChan As String)
    UpdateRTB WConsole.RTB, "*** " & szNick & " invited you to " & szChan & "!", mvarCLR_NOTICE
End Sub


Private Sub IRC_Ob_onNickJoin(ByVal szNick As String, ByVal szChan As String)
    If UCase$(szNick) = UCase$(NickName) Then
        'it twas me
        ChanNew szChan
        IRC_Ob.Send "NAMES " & szChan & vbCrLf
    Else
        'update list
        mvarChanWins.Item(szChan).nicklist.ListItems.Add 1, szNick, szNick
        UpdateRTB mvarChanWins.Item(szChan).chanrtb, "*** " & szNick & " just joined this channel!", mvarCLR_JOIN
    End If
    SCR_Ob.onChannelJoin szChan, szNick
End Sub

Private Sub IRC_Ob_onNickKick(ByVal szNick As String, ByVal szChan As String, ByVal szKicker As String, ByVal szComments As String)
            UpdateRTB mvarChanWins.Item(szChan).chanrtb, "*** " & szNick & " was kicked by " & szKicker & " (" & szComments & ")", mvarCLR_KICK
            mvarChanWins.Item(szChan).nicklist.ListItems.Remove szNick
            'KICK --we were kicked from a channel...damn
            If Trim(LCase(szNick)) = LCase(NickName) Then
                mvarChanWins.Item(szChan).Hide
                mvarChanWins.Remove szChan
                RaiseEvent onKillChannelWindow(szChan)
            End If
End Sub

Private Sub IRC_Ob_onNickPart(ByVal szNick As String, ByVal szChan As String)
    If Not LCase$(szNick) = LCase$(NickName) Then
        mvarChanWins.Item(szChan).nicklist.ListItems.Remove szNick
        UpdateRTB mvarChanWins.Item(szChan).chanrtb, "*** " & szNick & " just left this channel!", mvarCLR_PART
    End If
    SCR_Ob.onChannelPart szChan, szNick
End Sub


Private Sub IRC_Ob_onNickQuit(ByVal szNick As String)
    UpdateRTB WConsole.RTB, "*** " & szNick & " just quit irc!", mvarCLR_PART
    Dim num As Integer
        For num = 1 To mvarChanWins.count
            If existsInList(ChanWins.Item(num).nicklist, szNick) Then
                mvarChanWins.Item(num).nicklist.ListItems.Remove szNick
            End If
        Next num
End Sub


Private Sub IRC_Ob_onNotice(ByVal szSender As String, ByVal szMessage As String)
    UpdateRTB WConsole.RTB, "[" & szSender & "] " & szMessage, mvarCLR_NOTICE
End Sub

Private Sub IRC_Ob_onNotifyNicks(ByVal szNickList As String)
    RaiseEvent onUpdateNotifyWindow(szNickList)
    SCR_Ob.onNotifyList szNickList
End Sub

Private Sub IRC_Ob_onPrivMsg(ByVal szNick As String, ByVal szMessage As String)
    On Error GoTo here
    If IsObject(MessWins.Item(szNick)) Then
        UpdateRTB mvarMessWins.Item(szNick).messgrtb, "<" & szNick & ">   " & szMessage, mvarCLR_MESSAGE
    Else
here:
        MessNew szNick
        UpdateRTB mvarMessWins.Item(szNick).messgrtb, "<" & szNick & ">   " & szMessage, mvarCLR_MESSAGE
    End If
End Sub

Private Sub IRC_Ob_onServerMessage(ByVal szMessage As String)
    UpdateRTB WConsole.RTB, szMessage, mvarCLR_MESSAGE
End Sub


Private Sub IRC_Ob_onTopicChanged(ByVal szTopic As String, ByVal szChan As String, ByVal szNick As String)
    mvarChanWins.Item(szChan).Caption = "[" & szChan & "] " & szTopic
    UpdateRTB mvarChanWins.Item(szChan).chanrtb, "*** " & szNick & " set topic to " & szTopic, mvarCLR_MODE
End Sub
Public Sub Send(ByVal szLine As String)
    IRC_Ob.Send szLine
End Sub
Public Sub SendChatText(ByVal szLine As String, ByVal handle As Long)
    DCC_Ob.SendChatText szLine, handle
End Sub

Public Sub KillDcc(ByVal handle As Long)
    DCC_Ob.Destroy handle
End Sub
Public Sub GetVersion()
    RaiseEvent onQueryAppVersion("[" & App.Title & "] " & App.Major & "." & App.Minor & "." & App.Revision)
    RaiseEvent onQueryAppVersion(IRC_Ob.GetVersion())
    RaiseEvent onQueryAppVersion(BOT_Ob.GetVersion())
    RaiseEvent onQueryAppVersion(DCC_Ob.GetVersion())
    RaiseEvent onQueryAppVersion(ID_Ob.GetVersion())
    RaiseEvent onQueryAppVersion(HookMan.GetVersion())
End Sub

Public Sub DCCSendFile(ByVal szNick As String, ByVal szFileName As String, targfrm As Object)
    DCC_Ob.SendFile szNick, szFileName, FileLen(szFileName), 1024, targfrm
End Sub
Public Sub DCCSendchat(ByVal szNick As String, targfrm As Object)
    DCC_Ob.Sendchat szNick, targfrm
End Sub

Public Sub ResetScriptEngine()
    SCR_Ob.Reset
End Sub

Public Function ToggleScriptStatus() As Boolean
    SCR_Ob.Enabled = Not SCR_Ob.Enabled
    ToggleScriptStatus = SCR_Ob.Enabled
End Function
Public Sub SendFromScriptToWindow(ByVal szKey As String, ByVal szLine As String)
'Felmlee 13-Jun-98 21:46:59
    If Mid(szKey, 1, 1) = "#" Then
        UpdateRTB mvarChanWins.Item(szKey).chanrtb, szLine, mvarCLR_CTCP
    Else
        UpdateRTB mvarMessWins.Item(szKey).chanrtb, szLine, mvarCLR_CTCP
    End If
End Sub
