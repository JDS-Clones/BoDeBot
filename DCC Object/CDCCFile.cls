VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDCCFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarhWnd As Long 'local copy
Private mvarWndProc As Long
Private mvarFactory As CDCCFactory 'local copy
Private mvarFileName As String 'local copy
Private mvarFileSize As Long 'local copy
Private mvarChunkSize As Integer 'local copy
Private mvarHost As String 'local copy
Private mvarConnID As Long 'local copy
Private mvarFrm As Object 'local copy
Private mvarAddr As String 'local copy
Private mvarPort As Integer 'local copy

Private dwOperation As Integer

Const vbGetFile = 0
Const vbSendFile = 1

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



Public Property Let ConnID(ByVal vData As Long)
    mvarConnID = vData
End Property

Public Property Get ConnID() As Long
    ConnID = mvarConnID
End Property

Public Sub StopDCCFile()
End Sub

Public Property Let Host(ByVal vData As String)
    mvarHost = vData
End Property


Public Property Get Host() As String
    Host = mvarHost
End Property

Public Property Let chunksize(ByVal vData As Integer)
    mvarChunkSize = vData
End Property

Public Property Get chunksize() As Integer
    chunksize = mvarChunkSize
End Property

Public Property Let FileSize(ByVal vData As Long)
    mvarFileSize = vData
End Property

Public Property Get FileSize() As Long
    FileSize = mvarFileSize
End Property

Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Property Set Factory(ByVal vData As Object)
    Set mvarFactory = vData
End Property

Public Property Get Factory() As CDCCFactory
    Set Factory = mvarFactory
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

Private Sub AppendChunk(ByVal dwAction As Integer, Chunk() As Byte, Optional lpRecvd As Long)
    Dim i As Long
    Static filenum
    Select Case dwAction
    Case 0 'open file
        filenum = FreeFile(0)
        Open mvarFileName For Binary Access Write As #filenum
    Case 1 'appendchunk
        For i = 0 To lpRecvd - 1
            Put #filenum, , Chunk(i)
        Next i
    Case 2 'closefile
        Close #filenum
    End Select
End Sub
Private Function GetChunk(ByVal dwAction As Integer, ByRef buff() As Byte) As Byte
    Static filenum
    Select Case dwAction
    Case 0 'open file
        filenum = FreeFile(0)
        Open mvarFileName For Binary Access Read As #filenum
    Case 1 'getchunk
        Get #filenum, , buff
    Case 2 'closefile
        Close #filenum
    End Select
End Function
Public Sub DCCConnect()
    dwOperation = vbGetFile
    mvarFrm.Caption = "DCC Get: " & mvarHost & " (" & mvarAddr & ")"
    ConnectSock mvarAddr, mvarPort, 0, mvarhWnd, True 'connect to the client
End Sub

Public Sub DCCListen()
    Dim dwPort As Integer
    Dim addrlen&
    Dim ret&
    Dim sa As sockaddr
    dwOperation = vbSendFile
    mvarFrm.Caption = "DCC Send: " & mvarHost
    addrlen = sockaddr_size
    ret = getsockname(mvarFactory.lpIRC_SOCKET, sa, addrlen) 'get our ip address
    dwPort = ListenToFreePort(mvarConnID) 'start listening and get the port
    mvarPort = dwPort
    'CHAT chat 2155264017 1902
    Call SendData(mvarFactory.lpIRC_SOCKET, "PRIVMSG " & mvarHost & " :" & Chr(1) & "DCC SEND " & mvarFactory.ReturnFileName(mvarFileName) & " " & mvarFactory.EncodeIPAddr(getascip$(sa.sin_addr)) & " " & dwPort & " " & mvarFileSize & Chr(1) & vbCrLf)
End Sub
Public Function Hook(ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    Dim x As Long, a As String, i As Long, l&
    Dim ReadBuffer() As Byte
    ReDim ReadBuffer(mvarChunkSize) As Byte
    Dim SendD As String
    Dim saddr As sockaddr
    Dim temp As Long
    Static tot As Long
    Static CurrentLine As String
    Static was_delim As Boolean
    Select Case Msg
        Case 1025
            Select Case dwOperation
                '#Send File Form
                Case vbSendFile
                    Dim buffer() As Byte
                    ReDim buffer(mvarChunkSize)
                    Select Case lp
                        Case FD_READ
                            mvarConnID = wp
                            x = recv(mvarConnID, ReadBuffer(0), mvarChunkSize, 0)
                            If x > 0 Then
                                tot = tot + mvarChunkSize
                                If tot >= mvarFileSize Then
                                   'mvarFactory.RaiseUpdateProgress mvarChunkSize, mvarFrm
                                   mvarFactory.RaiseFileDone mvarHost, mvarFileName, mvarAddr, mvarFileSize, mvarFrm
                                   Call closesocket(mvarConnID)
                                   Call GetChunk(2, buffer) 'close file
                                Else
                                    mvarFactory.RaiseUpdateProgress mvarChunkSize, mvarFrm
                                    'send next chunk
                                    Call GetChunk(1, buffer)
                                    Call SendData(mvarConnID, buffer)
                                End If
                            End If
                        Case FD_CONNECT
                            mvarConnID = wp
                                mvarFactory.RaiseFileStart mvarHost, mvarFileName, mvarAddr, mvarFileSize, mvarChunkSize, mvarFrm, "SEND"
                        Case FD_ACCEPT
                            temp = mvarConnID
                            mvarConnID = accept(wp, saddr, sockaddr_size)
                            Call closesocket(temp)
                            tot = 0
                            mvarFactory.RaiseFileStart mvarHost, mvarFileName, mvarAddr, mvarFileSize, mvarChunkSize, mvarFrm, "SEND"
                            Call GetChunk(0, buffer) 'open file
                            Call GetChunk(1, buffer)
                            Call SendData(mvarConnID, buffer)
                        Case FD_CLOSE
                            Call closesocket(mvarConnID)
                            mvarFactory.RaiseFileDone mvarHost, mvarFileName, mvarAddr, mvarFileSize, mvarFrm
                            Call GetChunk(2, buffer) 'close file
                            Erase ReadBuffer
                            Erase buffer
                    End Select
                
                '#Get File Form
                Case vbGetFile
                    Select Case lp
                        Case FD_READ
                            mvarConnID = wp
                            x = recv(mvarConnID, ReadBuffer(0), mvarChunkSize, 0)
                            If x > 0 Then
                                tot = tot + x 'add another chunk to the total we have so far
                                If tot >= mvarFileSize Then
                                   l = htonl(tot)
                                    mvarFactory.RaiseFileDone mvarHost, mvarFileName, mvarAddr, mvarFileSize, mvarFrm
                                    Call SendData(mvarConnID, GetAck(l))
                                    'Append Chunk to file
                                    AppendChunk 1, ReadBuffer, x
                                    AppendChunk 2, ReadBuffer
                                Else
                                    mvarFactory.RaiseUpdateProgress x, mvarFrm
                                    AppendChunk 1, ReadBuffer, x
                                    l = htonl(tot)
                                    Call SendData(mvarConnID, GetAck(l))
                                End If
                            End If
                        Case FD_CONNECT
                            mvarConnID = wp
                            tot = 0
                            AppendChunk 0, ReadBuffer
                            mvarFactory.RaiseFileStart mvarHost, mvarFileName, mvarAddr, mvarFileSize, mvarChunkSize, mvarFrm, "GET"
                        Case FD_ACCEPT
                            'Never happens
                        Case FD_CLOSE
                            mvarConnID = wp
                            Call closesocket(mvarConnID)
                            AppendChunk 2, ReadBuffer
                            mvarFactory.RaiseFileDone mvarHost, mvarFileName, mvarAddr, mvarFileSize, mvarFrm
                            Erase ReadBuffer
                    End Select
            End Select
            Hook = 0
        Case Else
            Hook = mvarWndProc
    End Select
    
End Function
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

Private Sub Class_Terminate()
    Call closesocket(mvarConnID)
End Sub
Private Function GetAck(ByVal x As Long) As String
   Dim b As Byte
   If (x And &H80000000) Then
     b = ((x And &H7FFFFFFF) \ 16777216) Or &H80
   Else
     b = (x \ 16777216)
   End If
   GetAck = Chr$(x And &HFF) & Chr$((x \ 256) And &HFF) & Chr$((x \ 65536) And &HFF) & Chr$(b)
End Function
