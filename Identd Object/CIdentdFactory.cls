VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIdentdFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************************************
'   IdentD Server Component
'       Author: Nicholas J. Felmlee
'       Email: Nick@Felmlee.Com
'       Last Revision: 7/18/98
'
'       Internal Dependencies:
'           WskSock.bas
'           CIdentdFactory.cls
'
'       External Dependencies:
'           -none-
'
'**************************************************************

Option Explicit
'local variable(s) to hold property value(s)
Private mvarIdentName As String 'local copy
Private mvarIdentSystem As String 'local copy
Private mvarIdentPort As Integer 'local copy
Public Event onIdentRequest()
Private lpIDENT_SOCKET As Long
Private mvarhWnd As Long 'who we attach our hook to
Private mvarWndProc As Long
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


Public Sub Start()
    Call ListenForConnect(113, mvarhWnd)
End Sub

Public Property Let IdentPort(ByVal vData As Integer)
    mvarIdentPort = vData
End Property


Public Property Get IdentPort() As Integer
    IdentPort = mvarIdentPort
End Property



Public Property Let IdentSystem(ByVal vData As String)
    mvarIdentSystem = vData
End Property


Public Property Get IdentSystem() As String
    IdentSystem = mvarIdentSystem
End Property



Public Property Let IdentName(ByVal vData As String)
    mvarIdentName = vData
End Property


Public Property Get IdentName() As String
    IdentName = mvarIdentName
End Property



Public Function Hook(ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    Dim x As Long, a As String, i As Long
    Dim ReadBuffer(1) As Byte
    Static CurrentLine As String
    Dim rply As String, temp As Long
    Dim saddr As sockaddr
    
    Select Case msg
        Case 1025
            '#We have connected
            If lp = FD_CONNECT Then
                lpIDENT_SOCKET = wp
        
                
            '#Data recieved yippie
            ElseIf lp = FD_READ Then
                lpIDENT_SOCKET = wp
                x = recv(lpIDENT_SOCKET, ReadBuffer(0), 1, 0)
                If x > 0 Then
                    a = StrConv(ReadBuffer, vbUnicode)
                    a = CStr(Chr$(Asc(a)))
                    If Not (Asc(a) = 10 Or Asc(a) = 13 Or Asc(a) = 0) Then
                        CurrentLine = CurrentLine & a
                    Else
                        rply = CurrentLine & " : USERID : " & mvarIdentSystem & " : " & mvarIdentName
                        Call SendData(lpIDENT_SOCKET, rply)
                        Call closesocket(lpIDENT_SOCKET)
                        CurrentLine = ""
                    End If
                End If
        
            ElseIf lp = FD_ACCEPT Then
                Dim siz&
                siz& = sockaddr_size
                temp = wp
                lpIDENT_SOCKET = accept(temp, saddr, siz&)
                'Call closesocket(temp) 'kills listening socket... we want to keep responding
            
            ElseIf lp = FD_CLOSE Then
                lpIDENT_SOCKET = wp
                Call closesocket(lpIDENT_SOCKET)
            End If
            Hook = 0
        Case Else
            Hook = mvarWndProc
        End Select
End Function

Public Function GetVersion() As String
    GetVersion = "[" & App.Title & "] " & App.Major & "." & App.Minor & "." & App.Revision
End Function

