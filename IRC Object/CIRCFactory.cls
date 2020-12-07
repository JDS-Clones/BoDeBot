VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIRCFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************************************
'   IRC Object Component
'       Author: Nicholas J. Felmlee
'       Email: Nick@Felmlee.Com
'       Last Revision: 7/18/98
'
'       Internal Dependencies:
'           WskSock.bas
'           CIRCFactory.cls
'
'       External Dependencies:
'           -none-
'
'**************************************************************

Option Explicit
Public Event onBeginNickList(ByVal szChan As String, ByVal szNicks As String)
Public Event onEndNickList(ByVal szChan As String)
Public Event onError(ByVal lpErrCode As Long, ByVal szDescription As String)
'**Public Event onExitNick(ByVal szNick As String, ByVal szChan As String)
Public Event onServerMessage(ByVal szMessage As String)
Public Event onPrivMsg(ByVal szNick As String, ByVal szMessage As String)
Public Event onCTCP(ByVal szNick As String, ByVal szCommand As String, ByVal szExtras As String)
Public Event onChanMsg(ByVal szSender As String, ByVal szChan As String, ByVal szMessage As String)
Public Event onNotice(ByVal szSender As String, ByVal szMessage As String)
Public Event onNickChanged(ByVal szOriginalNick As String, ByVal szNewNick As String)
Public Event onNickInvite(ByVal szNick As String, ByVal szChan As String)
Public Event onNickKick(ByVal szNick As String, ByVal szChan As String, ByVal szKicker As String, ByVal szComments As String)
'***Public Event onUserMode(ByVal szNick As String, ByVal szMode As String, ByVal szChan As String)
Public Event onChanMode(ByVal szSetter As String, ByVal szChan As String, ByVal szMode As String, ByVal szUsers As String)
Public Event onTopicChanged(ByVal szTopic As String, ByVal szChannel As String, ByVal szNick As String)
Public Event onNickJoin(ByVal szNick As String, ByVal szChan As String)
Public Event onNickPart(ByVal szNick As String, ByVal szChan As String)
Public Event onNickQuit(ByVal szNick As String)
Public Event onNotifyNicks(ByVal szNickList As String)
Public Event onAwayStatus(ByVal szReply As String, ByVal dwCode As Integer)
Public Event onConnect(ByVal szHost As String)
Public Event onDisconnect(ByVal szHost As String)
Public Event onEndMOTD() 'fires on end of MOTD so we know that we are fully connected



'local variable(s) to hold property value(s)
Private mvarServerName As String 'local copy
Private mvarPort As Integer 'local copy
Private mvarNickName As String 'local copy
Private mvarRealName As String 'local copy
Private mvarUsername As String 'local copy
Public Connected As Boolean
Public lpIRC_SOCKET As Long
Public LocalHost As String
Private mvarRetries As Integer 'local copy
Private mvarhWnd As Long 'Required by MsgHook.dll. Who we attach our hook to
Private mvarWndProc As Long 'Required by MsgHook.dll.
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
Public Property Let Retries(ByVal vData As Integer)
    mvarRetries = vData
End Property

Public Property Get Retries() As Integer
    Retries = mvarRetries
End Property
Public Sub Send(ByVal szline As String)
    If SendData(lpIRC_SOCKET, szline) < 0 Then
        RaiseEvent onDisconnect(mvarServerName)
    End If
End Sub

Public Sub Disconnect()
    Call closesocket(lpIRC_SOCKET)
    Connected = False
End Sub

Public Sub Connect()
    ConnectSock mvarServerName, mvarPort, 0, mvarhWnd, True
End Sub


Public Function Hook(ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    Dim SendD As String
    Static m_ReadBuffer As String
    Dim lDummy&, a&, lPos&, AsyncError&
    Dim sTemp$, OutBuffer$, RecvBuffer$
    Dim lTemp&, lTmp&, sHost$, sTmp$, tmphost As HostEnt
Select Case msg
    Case 1025
        '#We have connected
        If lp = FD_CONNECT Then
            lpIRC_SOCKET = wp
            RaiseEvent onConnect(ServerName)
            SendD = "USER " & mvarUsername & " host.com " & AddrToIP(LocalHost) & " :" & mvarRealName & vbCrLf
            Call SendData(lpIRC_SOCKET, SendD)
            SendD = "NICK " & mvarNickName & vbCrLf
            Call SendData(lpIRC_SOCKET, SendD)
            SendD = "MODE " & mvarNickName & " +i" & vbCrLf
            Call SendData(lpIRC_SOCKET, SendD)
    
        '#Data recieved yippie
        ElseIf lp = FD_READ Then
                    lpIRC_SOCKET = wp
                    RecvBuffer = String$(4096, " ")
                    lDummy = recv(lpIRC_SOCKET, ByVal RecvBuffer, Len(RecvBuffer), 0)
                    If lDummy > 0 Then
                        m_ReadBuffer = m_ReadBuffer & Left$(RecvBuffer, lDummy)
                        While InStr(m_ReadBuffer, vbLf)
                            lPos = InStr(m_ReadBuffer, vbLf)
                                If lPos < Len(m_ReadBuffer) Then
                                    OutBuffer = Left$(m_ReadBuffer, lPos - 1)
                                    m_ReadBuffer = Mid$(m_ReadBuffer, lPos + 1)
                                        lPos = InStr(OutBuffer, vbCr)
                                        While lPos > 0
                                            If lPos < Len(OutBuffer) Then
                                                OutBuffer = Left$(OutBuffer, lPos - 1) & Mid$(OutBuffer, lPos + 1)
                                            Else
                                                OutBuffer = Left$(OutBuffer, lPos - 1)
                                            End If
                                            lPos = InStr(OutBuffer, vbCr)
                                        Wend
    '-------------------------------what to do after a linemode read
                                    ParseLine OutBuffer
    '----------------------------------------------------------------
                                Else
                                    OutBuffer = Left$(m_ReadBuffer, Len(m_ReadBuffer) - 1)
                                    m_ReadBuffer = ""
                                        lPos = InStr(OutBuffer, vbCr)
                                        While lPos > 0
                                            If lPos < Len(OutBuffer) Then
                                                OutBuffer = Left$(OutBuffer, lPos - 1) & Mid$(OutBuffer, lPos + 1)
                                            Else
                                                OutBuffer = Left$(OutBuffer, lPos - 1)
                                            End If
                                            lPos = InStr(OutBuffer, vbCr)
                                        Wend
    '-------------------------------what to do after a linemode read (same as above)
                                    ParseLine OutBuffer
    '-------------------------------------------------------------------------------
                                End If
                        Wend
                    End If
            
       '#closed
       Hook = 0
        ElseIf lp = FD_CLOSE Then
            Call closesocket(lpIRC_SOCKET)
            Connected = False
            RaiseEvent onDisconnect(ServerName)
        End If
    Case Else
        Hook = mvarWndProc
    End Select
End Function
Public Property Let Username(ByVal vData As String)
    mvarUsername = vData
End Property


Public Property Get Username() As String
    Username = mvarUsername
End Property



Public Property Let RealName(ByVal vData As String)
    mvarRealName = vData
End Property


Public Property Get RealName() As String
    RealName = mvarRealName
End Property



Public Property Let NickName(ByVal vData As String)
    mvarNickName = vData
End Property


Public Property Get NickName() As String
    NickName = mvarNickName
End Property

Public Property Let Port(ByVal vData As Integer)
    mvarPort = vData
End Property

Public Property Get Port() As Integer
    Port = mvarPort
End Property

Public Property Let ServerName(ByVal vData As String)
    mvarServerName = vData
End Property

Public Property Get ServerName() As String
    ServerName = mvarServerName
End Property

Private Sub processReply(rcode As Integer, ByVal szData$)
    Dim buff As String, temp As String, param1 As String, param2 As String, param3 As String
    '#Reserved numerics.
   'These numerics are not described above since they fall into one of
   'the following categories:
        '1. no longer in use;
        '2. reserved for future planned use;
        '3. in current use but are part of a non-generic 'feature' of the current IRC server.
        '209     RPL_TRACECLASS          217     RPL_STATSQLINE
        '231     RPL_SERVICEINFO         232     RPL_ENDOFSERVICES
        '233     RPL_SERVICE             234     RPL_SERVLIST
        '235     RPL_SERVLISTEND
        '316     RPL_WHOISCHANOP         361     RPL_KILLDONE
        '362     RPL_CLOSING             363     RPL_CLOSEEND
        '373     RPL_INFOSTART           384     RPL_MYPORTIS
        '466     ERR_YOUWILLBEBANNED     476     ERR_BADCHANMASK
        '492     ERR_NOSERVICEHOST
    Debug.Print "RPL " & rcode & vbCrLf
    Debug.Print "DATA= " & szData$
    Select Case rcode
    Case 300 'RPL_NONE
    Case 301 'RPL_AWAY
        Debug.Print "AWAY-> " & szData$
        RaiseEvent onAwayStatus(szData$, 301)
    Case 302 'RPL_USERHOST
        Debug.Print "USERHOST-> " & szData$
    Case 303 'RPL_ISON
        ':irc.vol.com 303 BoDeBoT :BoDePlOt BoDePlOt DrCert
        temp = Mid(szData$, 2)
        param1 = getNextToken(temp, ":")
        RaiseEvent onNotifyNicks(temp)
        'RaiseEvent onServerMessage("IS ON Reply -> " & temp)
    'NO 304
    Case 305 'RPL_UNAWAY
        ':irc.vol.com 305 BoDeBoT :You are no longer marked as being away
        temp = szData$
        Call getNextToken(temp, " ") 'irc.vol.com
        Call getNextToken(temp, " ") '305
        Call getNextToken(temp, " ") 'bodebot
        RaiseEvent onAwayStatus(Mid(temp, 2), 305)
    Case 306 'RPL_NOAWAY
        ':irc.vol.com 306 BoDeBoT :You have been marked as being away
        temp = szData$
        Call getNextToken(temp, " ") 'irc.vol.com
        Call getNextToken(temp, " ") '306
        Call getNextToken(temp, " ") 'bodebot
        RaiseEvent onAwayStatus(Mid(temp, 2), 306)
    'NO 307-310
    Case 311 'RPL_WHOISUSER
        ':irc.vol.com 311 BoDeBoT DrCert ~noone host-209-214-71-188.atl-n.bellsouth.net * :noone
        temp = szData$
        Call getNextToken(temp, " ") 'irc.vol.com
        Call getNextToken(temp, " ") '311
        Call getNextToken(temp, " ") 'bodebot
        param1 = getNextToken(temp, " ")
        RaiseEvent onServerMessage(param1 & " is " & temp)
    Case 312 'RPL_WHOISSERVER
        ':irc.vol.com 312 BoDeBoT DrCert irc.emory.edu :[170.140.4.6] Emory University
        temp = szData$
        Call getNextToken(temp, " ") 'irc.vol.com
        Call getNextToken(temp, " ") '312
        Call getNextToken(temp, " ") 'bodebot
        param1 = getNextToken(temp, " ")
        RaiseEvent onServerMessage(param1 & " using " & temp)
    Case 313 'RPL_WHOISOPERATOR
        Debug.Print "WHOISOPERATOR-> " & szData$
    Case 314 'RPL_WHOWASUSER
        Debug.Print "WHOWASUSER-> " & szData$
    Case 315 'RPL_ENDOFWHO
    'NO 316
    Case 317 'RPL_WHOISIDLE
        Debug.Print "WHOISIDLE-> " & szData$
    Case 318 'RPL_ENDOFWHOIS
        ':irc.vol.com 318 BoDeBoT drcert :End of /WHOIS list.
        temp = szData$
        Call getNextToken(temp, " ") 'irc.vol.com
        Call getNextToken(temp, " ") '311
        Call getNextToken(temp, " ") 'bodebot
        Call getNextToken(temp, " ") 'drcert
        RaiseEvent onServerMessage(Mid(temp, 2))
    Case 319 'RPL_WHOISCHANNELS
        Debug.Print "WHOISCHANNELS-> " & szData$
    'NO 320
    Case 321 'RPL_LISTSTART
    Case 322 'RPL_LIST
    Case 323 'RPLLISTEND
    Case 324 'RPL_CHANNELMODEIS
        Debug.Print "CHANNELMODEIS-> " & szData$
    'NO 325-330
    Case 331 'RPL_NOTOPIC
    Case 332 'RPL_TOPIC
        ':irc.vol.com 332 SMeLLMe #visualbasic :*** Wrexen has quit IRC (Licking some ass)
        buff = szData$
        temp = getNextToken(buff, " ") 'irc.vol.com
        temp = getNextToken(buff, " ") '332
        temp = getNextToken(buff, " ") ' SMeLLMe
        param1 = getNextToken(buff, " ") '#visualbasic
        param2 = Mid(buff, 2)
        RaiseEvent onTopicChanged(param2, param1, param1)
        
    'NO 333-340
    Case 341 'RPL_INVITING
        Debug.Print "INVITING-> " & szData$
    Case 342 'RPL_SUMMONING
    'NO 343-350
    Case 351 'RPL_VERSION
    Case 352 'RPL_WHOREPLY
        ':irc.vol.com 352 BoDeBoT #teenchat duh123 access-isdn1-22.oz.psu.edu irc2.sprynet.com Dingo1 H :5 Da  jEsTeR
        ':irc.vol.com 352 BoDeBoT #luthertech snag2 access-hnts1p9.hn.psu.edu irc.mcs.net Snag H@ :5 snag
        temp = szData$
        Call getNextToken(temp, " ") 'irc.vol.com
        Call getNextToken(temp, " ") '352
        Call getNextToken(temp, " ") 'bodebot
        RaiseEvent onServerMessage(temp)
    Case 353 'RPL_NAMREPLY
        ':irc.vol.com 353 SMeLLMe * #VisualBasic :SMeLLMe +wishdev FezzOr Larson Bver @CaiSSa OakSmoke ThundrBug T-2 Altec^ Victor RasterP Kanati aolsux AcSSDnID JeremyS Donkey Mephiston MaxS Weed7526 @reboot pentium Vv JBlaze Compe screwed @Moosebert MrParata @ShadowRam +Kiliman zipoff @sk8ball +EnEsch @MarC_BV @Shadey XLogility @Leahcim @SuiDo +English +MindRape @BenKo @Rebecca_ @Adriana O-3 @ML_
        buff = szData$
        temp = getNextToken(buff, " ") 'irc.vol.com
        temp = getNextToken(buff, " ") '353
        temp = getNextToken(buff, " ") 'smellme
        temp = getNextToken(buff, " ") ' *
        param1 = getNextToken(buff, " ") '#visualbasic
        param2 = Mid(buff, 2)
        RaiseEvent onBeginNickList(param1, param2)
        
    'NO 354-363
    Case 364 'RPL_LINKS
    Case 365 'RPL_ENDOFLINKS
    Case 366 'RPL_ENDOFNAMES
        ':irc.vol.com 366 SMeLLMe #visualbasic :End of /NAMES list.
        buff = szData$
        temp = getNextToken(buff, " ") 'irc.vol.com
        temp = getNextToken(buff, " ") '353
        temp = getNextToken(buff, " ") 'smellme
        param1 = getNextToken(buff, " ") '#visualbasic
        RaiseEvent onEndNickList(param1)
        
    Case 367 'RPL_BANLIST
        Debug.Print "BANLIST-> " & szData$
    Case 368 'RPL_ENDOFBANLIST
    Case 369 'RPL_ENDOFWHOWAS
    'NO 370
    Case 371 'RPL_INFO
        Debug.Print "INFO-> " & szData$
    Case 372 'RPL_MOTD
        ':irc.vol.com 372 BoDeBoT :-
        buff = szData$
        temp = getNextToken(buff, " ") 'irc
        temp = getNextToken(buff, " ") '372
        temp = getNextToken(buff, " ") 'bodebot
        RaiseEvent onServerMessage(Mid(buff, 2))
    'NO 373
    Case 374 'RPL_ENDOFINFO
    Case 375 'RPL_MOTDSTART
    Case 376 'RPL_ENDOFMOTD
        ':irc.vol.com 376 BoDeBoT :End of /MOTD command.
        buff = szData$
        temp = getNextToken(buff, " ") 'irc
        temp = getNextToken(buff, " ") '376
        temp = getNextToken(buff, " ") 'bodebot
        RaiseEvent onServerMessage(Mid(buff, 2))
        RaiseEvent onEndMOTD
    'NO 377-380
    Case 381 'RPL_YOUREOPER
        Debug.Print "YOUREOPER-> " & szData$
    Case 382 'RPL_REHASHING
    'NO 383-390
    Case 391 'RPL_TIME
    Case 392 'RPL_USERSSTART
    Case 393 'RPL_USERS
    Case 394 'RPL_ENDOFUSERS
    Case 395 'RPL_NOUSERS
    
    '200's start here
    Case 200 'RPL_TRACELINK
    Case 201 'RPL_TRACECONNECTING
    Case 202 'RPL_TRACEHANDSHAKE
    Case 203 'RPL_TRACEUNKNOWN
    Case 204 'RPL_TRACEOPERATOR
    Case 205 'RPL_TRACEUSER
    Case 206 'RPL_TRACESERVER
    'NO 207
    Case 208 'RPL_TRACENEWTYPE
    'NO 209-210
    Case 211 'RPL_STATSLINKINFO
    Case 212 'RPL_STATSCOMMANDS
    Case 213 'STATSCLINE
    Case 214 'STATSNLINE
    Case 215 'STATSILINE
    Case 216 'STATSKLINE
    'NO 217
    Case 218 'STATSYLINE
    Case 219 'ENDOFSTATS
    'NO 220
    Case 221 'RPL_UMODEIS
    'NO 222-240
    Case 241 'RPL_STATSLLINE
    Case 242 'RPL_STATSUPTIME
    Case 243 'RPL_STATSOLINE
    Case 244 'RPL_STATSHLINE
    'NO 245-250
    Case 251 'RPL_LUSERCLIENT
    Case 252 'RPL_LUSEROP
    Case 253 'RPL_LUSERUNKNOWN
    Case 254 'RPL_LUSERCHANNELS
    Case 255 'RPL_LUSERME
    Case 256 'RPL_ADMINME
    Case 257 'RPL_ADMINLOC1
    Case 258 'RPL_ADMINLOC2
    Case 259 'RPL_ADMINEMAIL
    Case 261 'RPL_TRACELOG
    
    End Select
End Sub
Private Sub processError(ecode As Integer)
    Debug.Print "ERR " & ecode & vbCrLf
    Dim szDesc As String
    Select Case ecode
        Case 401 'ERR_NOSUCHNICK
            szDesc = "No Such Nick"
        Case 402 'ERR_NOSUCHSERVER
            szDesc = "No Such Server"
        Case 403 'ERR_NOSUCHCHANNEL
            szDesc = "No Such Channel"
        Case 404 'ERR_CANNOTSENDTOCHANNEL
            szDesc = "Cannot Send To Channel"
        Case 405 'ERR_TOOMANYCHANNELS
            szDesc = "Too Many Channels"
        Case 406 'ERR_WASNOSUCHNICK
            szDesc = "Was No Such Nick"
        Case 407 'ERR_TOOMANYTARGETS
            szDesc = "Too Many Targets"
        'NO 408
        Case 409 'ERR_NOORIGN
            szDesc = "No Origin"
        'NO 410
        Case 411 'ERR_NORECIPIENT
            szDesc = "No Recipient"
        Case 412 'ERR_NOTEXTTOSEND
            szDesc = "No Text To Send"
        Case 413 'ERR_NOTOPLEVEL
            szDesc = "No Top Level"
        Case 414 'ERR_WILDTOPLEVEL
            szDesc = "Wild Top Level"
        'NO 415-420
        Case 421 'ERR_UNKNOWNCOMMAND
            szDesc = "Unknown Command"
        Case 422 'ERR_NOMOTD
            szDesc = "No MOTD"
        Case 423 'ERR_NOADMININFO
            szDesc = "No Admin Info"
        Case 424 'ERR_FILEERROR
            szDesc = "File Error"
        'NO 425-430
        Case 431 'ERR_NONICKNAMEGIVEN
            szDesc = "No Nickname Given"
        Case 432 'ERR_ERRONEUSNICKNAME
            szDesc = "Erronous Nickname"
        Case 433 'ERR_NICKNAMEINUSE
            szDesc = "Nickname In Use"
        'NO 434-435
        Case 436 'ERR_NICKCOLLISION
            szDesc = "Nick Collision"
        'NO 437-440
        Case 441 'ERR_USERNOTINCHANNEL
            szDesc = "User Not In Channel"
        Case 442 'ERR_NOTONCHANNEL
            szDesc = "Not On Channel"
        Case 443 'ERR_USERONCHANNEL
            szDesc = "User On Channel"
        Case 444 'ERR_NOLOGIN
            szDesc = "No Login"
        Case 445 'ERR_SUMMONDISABLED
            szDesc = "Summond Disabled"
        Case 446 'ERR_USERDISABLED
            szDesc = "User Disabled"
        'NO 447-450
        Case 451 'ERR_NOTREGISTERED
            szDesc = "Not Registered"
        'NO 452-460
        Case 461 'ERR_NEEDMOREPARAMS
            szDesc = "Need More Params"
        Case 462 'ERR_ALREADYREGISTERED
            szDesc = "Already Registered"
        Case 463 'ERR_NOPERMFORHOST
            szDesc = "No Permission For Your Host"
        Case 464 'ERR_PASSWDMISMATCH
            szDesc = "Password Mismatch"
        Case 465 'ERR_YOUREBANNEDCREEP
            szDesc = "You Are Banned Creep"
        'NO 466
        Case 467 'ERR_KEYSET
            szDesc = "Key Set"
        'NO 468-470
        Case 471 'ERR_CHANNELISFULL
            szDesc = "Channel Is Full"
        Case 472 'ERR_UNKNOWNMODE
            szDesc = "Unknown Mode"
        Case 473 'ERR_INVITEONLYCHAN
            szDesc = "Invite-Only Channel"
        Case 474 'ERR_BANNEDFROMCHAN
            szDesc = "You Are Banned From The Channel"
        Case 475 'ERR_BADCHANNELKEY
            szDesc = "Bad Channel Key"
        'NO 476-480
        Case 481 'ERR_NOPRIVILEGES
            szDesc = "No Privileges"
        Case 482 'ERR_CHANOPRIVSNEEDED
            szDesc = "Channel Op Privileges Needed"
        Case 483 'ERR_CANTKILLSERVER
            szDesc = "You Can't Kill the Server"
        'NO 484-490
        Case 491 'ERR_NOOPERHOST
            szDesc = "NOOPERHOST"
        'NO 492-500
        Case 501 'ERR_UMODEUNKNOWNFLAG
            szDesc = "User Mode Unknown Flag"
        Case 502 'ERR_USERSDONTMATCH
            szDesc = "Users Don't Match"
    End Select
    RaiseEvent onError(CLng(ecode), szDesc)
End Sub

Private Sub Class_Initialize()
    Dim retb As Boolean
    'StartWinsock ("")
    LocalHost = GetLocalHostName
End Sub

Private Function getNextToken(ByRef szTemp$, ByVal szTok$)
    Dim Y As Integer
    Y = InStr(szTemp$, szTok$)
    If Y <> 0 Then
        getNextToken = Left(szTemp$, Y - 1)
        szTemp$ = Mid(szTemp$, Y + 1)
    Else
        getNextToken = szTemp$
    End If
End Function

Private Function CountTokens(ByVal val As String) As Integer
    Dim count As Integer
    Dim x As Integer
    count = 0
    x = InStr(val, ",")
    Do Until x = 0
        val = Mid(val, x + 1)
        count = count + 1
        x = InStr(val, ",")
    Loop
    CountTokens = count + 1
End Function

Private Sub Class_Terminate()
    'EndWinsock
End Sub



Private Sub ParseLine(ByVal data As String)
    '************************************************
    ' Parses a line of input from the server
    '
    '************************************************
    Dim temp As String, buff As String, param1 As String, param2 As String
    data = Trim(data)
    Debug.Print data & vbCrLf
    '#Coded Response
    If (InStr(1, data, ":" & mvarServerName) = 1 Or InStr(1, Mid(data, 1, InStr(1, data, " ") - 1), "!") = 0) And (InStr(1, data, "PING") = 0) Then
        Dim szcode As Integer
        temp = Mid(data, InStr(1, data, " ") + 1)
        szcode = val(Mid(temp, 1, InStr(1, temp, " ") - 1))
        If szcode < 400 Then
            processReply szcode, data
        Else
            processError szcode
        End If
    '#server set our mode
    ElseIf (InStr(1, data, ":" & mvarNickName & " MODE")) = 1 Then
    
    '#Normal Messages
    ElseIf Mid(data, 1, 1) = ":" Then
        parseNormalMessages data
    
    '#NOTICE from server
    ElseIf InStr(1, data, "NOTICE") = 1 Then
        buff = data
        temp = getNextToken(buff, ":")
        RaiseEvent onNotice("Auth:", buff)
    '#PING from server
    ElseIf InStr(1, data, "PING") = 1 Then
        Call SendData(lpIRC_SOCKET, "PONG :" & mvarServerName & vbCrLf)
        'UpdateRTB WConsole.RTB, "***PONG PONG PONG", 1
    
    '#ERROR from server
    ElseIf InStr(1, data, "ERROR") = 1 Then
        MsgBox data, vbOKOnly, "BoDeBoT"
    '#something else we don't know about
    Else
        Debug.Print vbCrLf & "Unknown Line"
        'parseNormalMessages data
    End If
End Sub

Private Sub parseNormalMessages(ByVal szData As String)
    Dim data As String
    Dim buff As String
    data = szData
    Dim szNick As String
    Dim szCommand As String
    Dim szTarget As String
    Dim szParams As String
    Dim szcode As String
    Dim pos As Integer
    Dim temp As String
        'get sender
        szNick = Mid(getNextToken(data, "!"), 2)
        Debug.Print "Sender was: " & szNick & vbCrLf
        'get command
        temp = getNextToken(data, " ")
        szCommand = getNextToken(data, " ")
        Debug.Print "Command was: " & szCommand & vbCrLf
        'get object of command
        szTarget = getNextToken(data, " ")
        Debug.Print "Target was: " & szTarget & vbCrLf
        
        Select Case UCase$(szCommand)
            Case "PRIVMSG"
                szParams = Mid(data, 2)
                If InStr(1, szTarget, "#") <> 0 Then
                    RaiseEvent onChanMsg(szNick, LCase$(szTarget), szParams)
                Else
                    If Not Mid(szParams, 1, 1) = Chr(1) Then
                        RaiseEvent onPrivMsg(szNick, szParams)
                    Else
                        szParams = Mid(szParams, 2)
                        RaiseEvent onCTCP(szNick, getNextToken(szParams, " "), szParams)
                    End If
                End If
            Case "JOIN"
                RaiseEvent onNickJoin(szNick, LCase$(Mid(szTarget, 2)))
            Case "PART"
                RaiseEvent onNickPart(szNick, LCase$(szTarget))
            Case "QUIT"
                RaiseEvent onNickQuit(szNick)
            Case "KICK"
                RaiseEvent onNickKick(getNextToken(data, " "), LCase$(szTarget), szNick, Mid(data, 2))
            Case "TOPIC"
                RaiseEvent onTopicChanged(Mid(data, 2), LCase$(szTarget), szNick)
            Case "NOTICE"
                RaiseEvent onNotice(szNick, "[" & szTarget & "] " & Mid(data, 2))
            Case "MODE"
                RaiseEvent onChanMode(szNick, szTarget, getNextToken(data, " "), data)
            Case "NICK"
                RaiseEvent onNickChanged(szNick, Mid(szTarget, 2))
            Case "INVITE"
                RaiseEvent onNickInvite(szNick, Mid(data, 2))
        End Select
End Sub

Public Function GetVersion() As String
    GetVersion = "[" & App.Title & "] " & App.Major & "." & App.Minor & "." & App.Revision
End Function
