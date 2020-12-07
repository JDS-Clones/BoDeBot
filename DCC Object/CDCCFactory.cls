VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDCCFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************************************
'   DCC Object Component
'       Author: Nicholas J. Felmlee
'       Email: Nick@Felmlee.Com
'       Last Revision: 07/18/98
'
'       Internal Dependencies:
'           WskSock.bas
'           CDCCFactory.cls
'           CDCCConn.cls
'           CDCCFile.cls
'
'       External Dependencies:
'           -none-
'
'**************************************************************

Option Explicit
Public Event onDCCChatClose(ByVal szNick As String, targfrm As Object)
Public Event onDCCChatOpen(ByVal szNick As String, targfrm As Object)
Public Event onDCCFileDone(ByVal szNick As String, ByVal szFileName As String, ByVal szHostAddr As String, ByVal lpFileSize As Long, targfrm As Object)
Public Event onDCCFileStart(ByVal szNick As String, ByVal szFileName As String, ByVal szHostAddr As String, ByVal lpFileSize As Long, ByVal chunksize As Long, targfrm As Object, ByVal szOp As String)
Public Event onUpdateProgress(ByVal lpChunk As Long, targfrm As Object)
Public Event onWindowOutput(ByVal szLine As String, targfrm As Object)

Private mvardwTimeOut As Integer 'time out value in seconds
Private mvarlpIRC_SOCKET As Long 'handle to our irc connection
Private DccColl As New Collection 'collection of instances of dccconns
'uses the form handle as the key

Public Event onAddDCCConn(Ob As Object)
Public Event onDelDCCConn(Ob As Object)

Private Parent As CDCCFactory


Public Sub Destroy(ByVal hWnd As Long)
'THIS DESTROYS AN OBJECT WHEN
'THE USER KILLS ONE OF OUR DCC WINDOWS.
    Dim szKey As String
    szKey = CStr(hWnd)
    DelConn DccColl.Item(szKey)
End Sub

Public Property Let lpIRC_SOCKET(ByVal vData As Long)
    mvarlpIRC_SOCKET = vData
End Property
Public Property Get lpIRC_SOCKET() As Long
    lpIRC_SOCKET = mvarlpIRC_SOCKET
End Property
Public Property Let dwTimeOut(ByVal vData As Integer)
    mvardwTimeOut = vData
End Property
Public Property Get dwTimeOut() As Integer
    dwTimeOut = mvardwTimeOut
End Property

Friend Function DecodeLongIPAddr(ByVal LongIPAddr As String) As String
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
Friend Function EncodeIPAddr(IPAddr As String) As String
    Dim DStart, EIP, DStop, ThisByte, HiWord!, LoWord!, LongIP
    DStart = 1
    EIP = ""
    Do
       DStop = InStr(DStart, IPAddr & ".", ".")
       ThisByte = Hex(Val(Mid$(IPAddr & ".", DStart, DStop - DStart)))
       EIP = EIP & IIf(Len(ThisByte) = 1, "0" & ThisByte, ThisByte)
       DStart = DStop + 1
    Loop Until DStart >= Len(IPAddr & ".")
    HiWord! = Val("&H" & Mid(EIP, 1, 2)) * 256! + Val("&H" & Mid(EIP, 3, 2))
    LoWord! = Val("&H" & Mid(EIP, 5, 2)) * 256! + Val("&H" & Mid(EIP, 7, 2))
    LongIP = HiWord! * 65536 + LoWord!
    EncodeIPAddr = Trim$(Str$(LongIP))
End Function


Friend Function ReturnFileName(ByVal vData As String) As String
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


Friend Sub AddConn(Ob As Object)
'ADD A CONNECTION TO OUR COLLECTION
'ALSO TAKES CARE OF HOOKING IT
    Dim szKey As String
    szKey = CStr(Ob.hWnd)
    RaiseEvent onAddDCCConn(Ob)
    DccColl.Add Ob, szKey
End Sub
Friend Sub DelConn(Ob As Object)
'REMOVE A CONNECTION FROM OUR COLLECTION
'ASLO TAKES CARE OF UNHOOKING IT
    Dim szKey As String
    szKey = CStr(Ob.hWnd)
    RaiseEvent onDelDCCConn(Ob)
    DccColl.Remove szKey
    Set Ob = Nothing
End Sub

Public Sub GetFile(ByVal szNick As String, ByVal szFileName As String, ByVal lpFileSize As Long, ByVal szHost As String, ByVal dwPort As Integer, ByVal lpChunkSize As Long, targfrm As Object)
'ACK A DCC FILE SEND REQUEST
    'create new Ob
    Dim gfile As CDCCFile
    Set gfile = New CDCCFile
    'set some properties
    gfile.hWnd = targfrm.hWnd
    Set gfile.Frm = targfrm
    Set gfile.Factory = Parent
    'add the Ob to the hook manager
    AddConn gfile
    gfile.FileName = szFileName
    gfile.FileSize = lpFileSize
    gfile.Host = szNick
    gfile.Addr = szHost
    gfile.Port = dwPort
    gfile.chunksize = lpChunkSize
    'start the transfer
    gfile.DCCConnect
End Sub

Public Sub SendFile(ByVal szNick As String, ByVal szFileName As String, ByVal lpFileSize As Long, ByVal lpChunkSize As Long, targfrm As Object)
'SEND A DCC FILE TRANSFER REQUEST
    'create new Ob
    Dim sfile As CDCCFile
    Set sfile = New CDCCFile
    'set some properties
    sfile.hWnd = targfrm.hWnd
    Set sfile.Frm = targfrm
    Set sfile.Factory = Parent
    'add the Ob to the hook manager
    AddConn sfile
    sfile.FileName = szFileName
    sfile.FileSize = lpFileSize
    sfile.Host = szNick
    If lpChunkSize > lpFileSize Then
        lpChunkSize = lpFileSize / 2
    End If
    sfile.chunksize = lpChunkSize
    'start the transfer
    sfile.DCCListen
End Sub

Public Sub GetChat(ByVal szNick As String, ByVal szHost As String, ByVal dwPort As Integer, targfrm As Object)
'ACK A DCC CHAT REQUEST
    'create new Ob
    Dim gchat As CDCCChat
    Set gchat = New CDCCChat
    'set some properties
    gchat.hWnd = targfrm.hWnd
    Set gchat.Frm = targfrm
    Set gchat.Factory = Parent
    'add the Ob to the hook manager
    AddConn gchat
    gchat.szNick = szNick
    gchat.Addr = szHost
    gchat.Port = dwPort
    'start chat
    gchat.DCCConnect
End Sub

Public Sub SendChat(ByVal szNick As String, targfrm As Object)
'SEND A DCC CHAT REQUEST
    Dim schat As CDCCChat
    Set schat = New CDCCChat
    'set some properties
    schat.hWnd = targfrm.hWnd
    Set schat.Frm = targfrm
    Set schat.Factory = Parent
    'add the Ob to the hook manager
    AddConn schat
    schat.szNick = szNick
    'start chat
    schat.DCCListen
End Sub

Private Sub Class_Initialize()
    Set Parent = Me
    Set DccColl = New Collection
End Sub
Friend Sub RaiseChatOpen(ByVal szNick As String, targfrm As Object)
'THROWS AN EVENT THAT TELLS US A DCC CHAT HAS STARTED
'AND WHAT FORM SHOULD REFLECT THAT
    RaiseEvent onDCCChatOpen(szNick, targfrm)
End Sub
Friend Sub RaiseChatClose(ByVal szNick As String, targfrm As Object)
'THROWS AN EVENT THAT TELLS US A DCC CHAT HAS CLOSED
'AND WHAT FORM SHOULD INDICATE IT
    RaiseEvent onDCCChatClose(szNick, targfrm)
End Sub
Friend Sub RaiseFileStart(ByVal szNick As String, ByVal szFileName As String, ByVal szHost As String, ByVal lpSize As Long, ByVal chunksize As Long, targfrm As Object, ByVal szOp As String)
'THROWS AN EVENT THAT TELLS US OUR DCC TRANSFER HAS STARTED
'AND WHAT FORM SHOULD UPDATE ITS STATUS
    RaiseEvent onDCCFileStart(szNick, szFileName, szHost, lpSize, chunksize, targfrm, szOp)
End Sub
Friend Sub RaiseFileDone(ByVal szNick As String, ByVal szFileName As String, ByVal szHost As String, ByVal lpSize As Long, targfrm As Object)
'THROWS AN EVENT THAT TELLS US ONE OF OUR FILE DCCS IS COMPLETE
'AND WHICH FORM SHOULD SHOW "DONE"
    RaiseEvent onDCCFileDone(szNick, szFileName, szHost, lpSize, targfrm)
End Sub
Friend Sub RaiseUpdateProgress(ByVal lpChunk As Long, targfrm As Object)
'THROWS AN EVENT THAT TELLS US THAT WE RECIEVED ANOTHER FILE CHUNK AND WHAT
'FORM WE SHOULD UPDATE A PROGRESS INDICATOR ON.
    RaiseEvent onUpdateProgress(lpChunk, targfrm)
End Sub
Friend Sub RaiseOutputWindow(ByVal szLine As String, targfrm As Object)
'THROWS AN EVENT TO TELL US THAT WE HAVE DCC CHAT INPUT AND
'TELLS US WHAT FORM WE SHOULD DISPLAY IT ON.
    RaiseEvent onWindowOutput(szLine, targfrm)
End Sub


Public Sub SendChatText(ByVal szLine As String, ByVal hWnd As Long)
'SENDS OUR USER INPUT IN THE DCC CHAT CONNECTION
    Dim szKey As String
    szKey = CStr(hWnd)
    'send text thru our chat connection
    Call SendData(DccColl.Item(szKey).ConnID, szLine)
End Sub
Public Function GetVersion() As String
    GetVersion = "[" & App.Title & "] " & App.Major & "." & App.Minor & "." & App.Revision
End Function

