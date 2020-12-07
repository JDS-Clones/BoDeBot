VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDCCChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarFactory As CDCCFactory 'local copy
Private mvarszNick As String 'local copy
Private mvarConnID As Long 'local copy
Private mvarhWnd As Long 'local copy
Private mvarWndProc As Long
Private mvarFrm As Object 'local copy
Private mvarAddr As String 'local copy
Private mvarPort As Integer 'local copy

Private dwOperation As Integer
Const vbGetChat = 0
Const vbSendChat = 1
Public Property Let Port(ByVal vData As Integer)
    mvarPort = vData
End Property

Public Property Get Port() As Integer
    Port = mvarPort
End Property

Public Property Let Addr(ByVal vData As String)
    mvarAddr = vData
End Property

Public Property Get Addr() As String
    Addr = mvarAddr
End Property

Public Property Set Frm(ByVal vData As Object)
    Set mvarFrm = vData
End Property


Public Property Get Frm() As Object
    Set Frm = mvarFrm
End Property



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

Public Property Let ConnID(ByVal vData As Long)
    mvarConnID = vData
End Property

Public Property Get ConnID() As Long
    ConnID = mvarConnID
End Property

Public Sub StopDCCChat()
End Sub

Public Property Let szNick(ByVal vData As String)
    mvarszNick = vData
End Property

Public Property Get szNick() As String
    szNick = mvarszNick
End Property

Public Property Set Factory(ByVal vData As Object)
    Set mvarFactory = vData
End Property

Public Property Get Factory() As CDCCFactory
    Set Factory = mvarFactory
End Property

Private Sub ProcessChatInput(ByVal szLine As String)
    mvarFactory.RaiseOutputWindow "<" & szNick & "> " & szLine, mvarFrm
End Sub

Public Sub DCCConnect()
    dwOperation = vbGetChat
    mvarFrm.Caption = "DCC Chat: " & mvarszNick & " (" & mvarAddr & ")"
    ConnectSock mvarAddr, mvarPort, 0, mvarhWnd, True 'connect to the client
End Sub
Public Sub DCCListen()
    Dim dwPort As Integer
    Dim addrlen&
    Dim ret&
    Dim sa As sockaddr
    dwOperation = vbSendChat
    mvarFrm.Caption = "DCC Chat: " & mvarszNick
    addrlen = sockaddr_size
    ret = getsockname(mvarFactory.lpIRC_SOCKET, sa, addrlen) 'get our ip address
    dwPort = ListenToFreePort(mvarConnID) 'start listening and get the port
    mvarPort = dwPort
    'CHAT chat 2155264017 1902
    Call SendData(mvarFactory.lpIRC_SOCKET, "PRIVMSG " & mvarszNick & " :" & Chr(1) & "DCC CHAT chat " & mvarFactory.EncodeIPAddr(getascip$(sa.sin_addr)) & " " & dwPort & Chr(1) & vbCrLf)
End Sub
Private Function ListenToFreePort(ByRef MySock As Long) As Integer
    Dim sockin As sockaddr
    Dim addrlen&, s, Dummy, SelectOps
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
    ListenToFreePort = htons(sadd.sin_port) '+ 32767
End Function
Public Function Hook(ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    Dim x As Long, a As String, i As Long
    Dim ReadBuffer() As Byte
    ReDim ReadBuffer(1)
    Dim SendD As String
    Dim saddr As sockaddr
    Dim temp As Long
    Static tot As Long
    Static CurrentLine As String
    Static was_delim As Boolean
    
    Select Case Msg
        Case 1025
        '#Chat Form
        'Case vbGetChat ', vbGetChat
            Select Case lp
                Case FD_READ
                    mvarConnID = wp
                    x = recv(wp, ReadBuffer(0), 1, 0)
                    If x > 0 Then
                        a = StrConv(ReadBuffer, vbUnicode)
                        a = CStr(Chr$(Asc(a)))
                        If Not ((a = Chr$(10)) Or (a = Chr$(13)) Or (a = (Chr$(10) & Chr$(13)))) Then
                            CurrentLine = CurrentLine & a
                            was_delim = False
                        Else
                            If Not was_delim Then
                                ProcessChatInput CurrentLine 'do something with the input
                                CurrentLine = ""
                                was_delim = True
                            End If
                        End If
                    End If
                Case FD_CONNECT
                    mvarConnID = wp
                    mvarFactory.RaiseChatOpen mvarszNick, mvarFrm
                Case FD_ACCEPT
                    temp = wp
                    mvarConnID = accept(wp, saddr, sockaddr_size)
                    Call closesocket(temp) 'close listener socket
                    mvarFactory.RaiseChatOpen mvarszNick, mvarFrm
                Case FD_CLOSE
                    Call closesocket(wp)
                    mvarFactory.RaiseChatClose mvarszNick, mvarFrm
            End Select
            Hook = 0
        Case Else
        Hook = mvarWndProc
    End Select
    
    
End Function

Private Sub Class_Terminate()
    Call closesocket(mvarConnID)
End Sub
